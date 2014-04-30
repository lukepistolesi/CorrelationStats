require 'thread_safe'

class CorrelationEngine

  attr_reader :configuration, :simple_stats, :data_samples_count,
              :rule_combinations, :probabilities, :conditionals

  def initialize(configuration)
    @configuration = configuration
    @simple_stats = { }
    configuration.rules.each_pair do |column_name, rules|
      @simple_stats[column_name] = { }
      rules.each_pair do |label, rule|
        #@simple_stats[column_name][label] = { }
        @simple_stats[column_name][label] = []
      end
    end
  end

  # Empty cells not supported
  def compute_simple_stats(workbook)
    data_sheet = workbook[configuration.data_sheet_name]
    puts "Data sheet loaded for '#{configuration.data_sheet_name}': computing simple stats..."

    current_column = configuration.data_column_idx
    header_string = Helper.cell_value data_sheet, configuration.data_row_idx, current_column

    start_row_index = current_row = find_data_series_start data_sheet, configuration.data_row_idx + 1, current_column
    columns_count = find_colums_extent data_sheet, configuration.data_row_idx, configuration.data_column_idx
    @data_samples_count = find_data_series_extent data_sheet, start_row_index, configuration.data_column_idx, columns_count

    puts "Data extent indexes row #{start_row_index} -- col #{configuration.data_column_idx}  to  row #{start_row_index + @data_samples_count} -- col #{configuration.data_column_idx + columns_count}"

    (configuration.data_column_idx..configuration.data_column_idx + columns_count).each_with_index do |col_idx|
      header_string = Helper.cell_value data_sheet, configuration.data_row_idx, col_idx
      puts "Processing column #{header_string}"

      (start_row_index..start_row_index + @data_samples_count - 1).each_with_index do |row_idx|
        rules = configuration.rules[header_string]

        cell_value = Helper.cell_value data_sheet, row_idx, col_idx
        compute_simple_stats_for_cell cell_value, row_idx, header_string, rules unless rules.nil?
      end

      header_string = Helper.cell_value data_sheet, configuration.data_row_idx, current_column
    end

    puts 'Simple stats computed'
  end

  def compute_bayes_correlations(workbook)
    puts 'Computing Bayes correlations'
    puts 'Step 1) Building rule combinations'
    start = Time.now
    @rule_combinations = build_rule_combinations(@configuration.rules).sort
    puts 'Step 2) Computing probabilities'
    compute_probabilities @rule_combinations
    puts 'Step 3) Computing conditional probabilites'
    compute_conditional_probabilities @rule_combinations, @probabilities
    total_seconds = Time.now - start
    puts "Bayes correlations computed in #{total_seconds} seconds (#{total_seconds.to_f/60.0} minutes)"
  end

  def to_s
    string = 'Simple Stats'
    @simple_stats.each_pair do |column_name, rules_stats|
      string.concat "\n  Stats for #{column_name}"
      rules_stats.each_pair do |label, stats|
        perc = stats.size.percent_of @data_samples_count
        string.concat "\n\t#{label} => #{stats.size} out of #{@data_samples_count} (#{perc})"
      end
    end

    string.concat "\n\n#{@rule_combinations.size} combinations built"

    string.concat "\n\nConditional probabilities"
    @conditionals.sort_by { |comb, results| results[:probability] }.each do |key_val|
      results = key_val[1]
      posterior = results[:posterior]
      probability = results[:probability].round(2) * 100
      unless posterior.nil? || probability <= 50.0 || posterior.count('&') > 1
        string.concat "\nP(#{results[:prior]}|#{posterior}) = #{probability}%"
      end
    end

    string
  end


  private

  def find_colums_extent(data_sheet, headers_row, start_col)
    cur_col = start_col
    while !Helper.cell_value(data_sheet, headers_row, cur_col).nil?
      cur_col = cur_col + 1
    end
    (cur_col - start_col) - 1
  end

  def find_data_series_start(data_sheet, row, col)
    begin
      cell_value = Helper.cell_value data_sheet, row, col
    end while cell_value.nil? && (row = row + 1)
    row
  end

  def find_data_series_extent(data_sheet, row, col, columns_count)
    cur_row = row
    begin
      empty_cells = 0
      (col..col + columns_count).each_with_index do |cur_col|
        value = Helper.cell_value(data_sheet, cur_row, cur_col)
        empty_cells = empty_cells + 1 if value.nil?
      end
      cur_row = cur_row + 1
    end while empty_cells < columns_count
    cur_row - row - 1
  end

  def compute_simple_stats_for_cell(cell_value, cell_index, header_string, rules_collection)
    column_stats = @simple_stats[header_string]
    rules_collection.each_pair do |label, rule|
      column_stats[label] << cell_index if !cell_value.nil? && rule[:proc].call(cell_value)
    end
  end

  def build_rule_combinations(rules_collection)
    all_combinations = ThreadSafe::Array.new
    threads = []
    start = Time.now
    rules_collection.each_with_index do |(root_column, rules), index|
      threads << Thread.new do
        other_columns = configuration.rules.keys - [root_column]
        rules.each_pair do |root_label, rule|
          combinations = ["#{root_column}:#{root_label}"]
          other_columns.each do |column_name|
            new_combinations = []
            configuration.rules[column_name].keys.each do |current_label|
              combinations.each { |comb| new_combinations << comb + "&#{column_name}:#{current_label}" }
            end
            combinations.concat new_combinations
          end
          all_combinations.concat combinations
        end
      end
    end

    threads.each { |thread| thread.join }

    total_time = Time.now - start
    puts "Total time to compute rule combinations is #{total_time} seconds (#{total_time.to_f/60.0} minutes)"
    all_combinations
  end

  def compute_probabilities(rule_combinations)
    puts "Combinations of rules to compute: #{rule_combinations.size.format_thousands}"
    @probabilities = {}
    @probabilities_cache_not_hit = 0
    intersection = []
    first_rule = last_rules = nil
    done = 0
    total_rules = rule_combinations.size
    batch_size = total_rules / 10
    start = Time.now

    rule_combinations.each do |combination|
      inter_idx = combination.index('&')
      first_rule = combination[0..inter_idx.to_i-1]
      intersection = compute_intersection_and_prob(first_rule) if @probabilities[first_rule].nil?
      if (last_rules = combination[first_rule.size + 1.. -1])
        #intersection = Helper.array_intersection intersection, compute_intersection_and_prob(last_rules)
        intersection = intersection & compute_intersection_and_prob(last_rules)
      end
      @probabilities[combination] ||=
        { intersection: intersection, probability: intersection.size.to_f / @data_samples_count.to_f }

      done = done + 1
      if done % batch_size == 0
        puts "Combination batch computed: #{done.to_f.percent_of(total_rules).round 2}"
      end
    end

    total_time = Time.now - start
    puts "Total time to compute simple probabilities is #{total_time} seconds (#{total_time.to_f/60.0} minutes)"
    puts "Cached probabilities hits #{(total_rules - @probabilities_cache_not_hit).percent_of(total_rules).round 2}%"
  end

  def compute_intersection_and_prob(combination)
    return @probabilities[combination][:intersection] if @probabilities[combination]
    @probabilities_cache_not_hit = @probabilities_cache_not_hit + 1
    intersection = nil
    combination.split('&').each do |col_and_rule|
      col, rule_label = col_and_rule.split ':'
      rule_idxs = @simple_stats[col][rule_label]
      intersection = intersection.nil? ? rule_idxs : (intersection & rule_idxs)
      #intersection = intersection.nil? ? rule_idxs : Helper.array_intersection(intersection, rule_idxs)
      break if intersection.empty?
    end
    @probabilities[combination] =
      { intersection: intersection, probability: intersection.size.to_f / @data_samples_count.to_f }
    intersection
  end

  def compute_conditional_probabilities(rule_combinations, probabilities)
    total_rules = rule_combinations.size
    batches = 12
    batch_size = (total_rules / batches) + 1
    done = 0

    start = Time.now
    @conditionals = ThreadSafe::Hash.new
    threads = []
    (0..batches-1).each_with_index do |idx|
      threads << Thread.new do
        local_conditionals = {}
        rule_combinations[idx*batch_size..(idx+1)*batch_size - 1].each do |combination|
          first_rule, second_rule = Helper.split_combo_rule_in_head_and_tail combination

          conditional_prob = 0.0
          numerator = probabilities[combination][:probability]

          if numerator != 0.0
            denumerator = probabilities[second_rule].nil? ? 1.0 : probabilities[second_rule][:probability]
            conditional_prob = numerator / denumerator
          end

          local_conditionals[combination] =
            { prior: first_rule, posterior: second_rule, probability: conditional_prob}

          done = done + 1
          if done % batch_size == 0
            puts "Combination batch computed: #{done.to_f.percent_of(total_rules).round 2}"
          end
        end
        @conditionals.merge! local_conditionals
      end
    end

    threads.each { |thread| thread.join }

    total_time = Time.now - start
    puts "Tot comb #{rule_combinations.size} -- computed #{@conditionals.size}"
    puts "Total time to compute conditional probabilities is #{total_time} seconds (#{total_time.to_f/60.0} minutes)"
  end
end
