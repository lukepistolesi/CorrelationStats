class CorrelationEngine

  attr_reader :configuration, :simple_stats, :data_samples_count,
              :rule_combinations, :probabilities, :conditionals

  def initialize(configuration)
    @configuration = configuration
    @simple_stats = { }
    configuration.rules.each_pair do |column_name, rules|
      @simple_stats[column_name] = { }
      rules.each_pair do |label, rule|
        @simple_stats[column_name][label] = { }
      end
    end
  end

  # Empty cells not supported
  def compute_simple_stats(workbook)
    data_sheet = workbook[configuration.data_sheet_name]
    puts "Data sheet loaded for '#{configuration.data_sheet_name}': computing simple stats..."

    current_column = configuration.data_column_idx
    header_string = Helper.cell_value data_sheet, configuration.data_row_idx, current_column

    while !header_string.nil? do
      puts "Processing column #{header_string}"

      start_row_index = current_row = find_data_series_start data_sheet, configuration.data_row_idx + 1, current_column
      cell_value = Helper.cell_value data_sheet, current_row, current_column

      while !cell_value.nil? do
        rules = configuration.rules[header_string]

        compute_simple_stats_for_cell cell_value, current_row, header_string, rules unless rules.nil?

        current_row = current_row + 1
        cell_value = Helper.cell_value data_sheet, current_row, current_column
      end

      current_column = current_column + 1
      header_string = Helper.cell_value data_sheet, configuration.data_row_idx, current_column
    end

    @data_samples_count = current_row - start_row_index
    puts 'Simple stats computed'
  end

  def compute_bayes_correlations(workbook)
    puts 'Computing Bayes correlations'
    puts 'Step 1) Building rule combinations'
    @rule_combinations = build_rule_combinations(@configuration.rules).sort
    puts 'Step 2) Computing probabilities'
    compute_probabilities @rule_combinations
    puts 'Step 3) Computing conditional probabilites'
    compute_conditional_probabilities @rule_combinations, @probabilities
    puts 'Bayes correlations computed'
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
      unless posterior.nil? || probability == 0.0
        string.concat "\nP(#{results[:prior]}|#{posterior}) = #{probability}%"
      end
    end

    string
  end


  private

  def find_data_series_start(data_sheet, row, col)
    begin
      cell_value = Helper.cell_value data_sheet, row, col
    end while cell_value.nil? && (row = row + 1)
    row
  end

  def compute_simple_stats_for_cell(cell_value, cell_index, header_string, rules_collection)
    column_stats = @simple_stats[header_string]

    rules_collection.each_pair do |label, rule|
      column_stats[label][cell_index] = cell_value if rule[:proc].call cell_value
    end
  end

  def build_rule_combinations(rules_collection)
    all_combinations = []
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
              combinations.each { |comb| new_combinations << comb + "∩#{column_name}:#{current_label}" }
            end
            combinations.concat new_combinations
          end
          all_combinations.concat combinations
        end
      end
    end

    threads.each { |thread| thread.join }

    total_time = ((Time.now - start).to_f / 60.0).round 2
    puts "Total time to compute rule combinations is #{total_time} minutes"
    all_combinations
  end

  def compute_probabilities(rule_combinations)
    @probabilities = {}
    rule_combinations.each do |combination|
      intersection = nil
      combination.split('∩').each do |col_and_rule|
        col, rule_label = col_and_rule.split ':'
        rule_idxs = @simple_stats[col][rule_label].keys
        intersection = intersection.nil? ? rule_idxs : (intersection & rule_idxs)
        break if intersection.empty?
      end
      @probabilities[combination] =
        { intersection: intersection, probability: intersection.size.to_f / @data_samples_count.to_f }
    end
  end

  def compute_conditional_probabilities(rule_combinations, probabilities)
    @conditionals = {}
    rule_combinations.each do |combination|
      first_rule, second_rule = Helper.split_combo_rule_in_head_and_tail combination

      conditional_prob = 0.0
      numerator = probabilities[combination][:probability]

      if numerator != 0.0
        denumerator = probabilities[second_rule].nil? ? 1.0 : probabilities[second_rule][:probability]
        conditional_prob = numerator / denumerator
      end

      @conditionals[combination] =
        { prior: first_rule, posterior: second_rule, probability: conditional_prob}
    end
  end
end
