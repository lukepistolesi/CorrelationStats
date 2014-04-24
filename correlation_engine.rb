class CorrelationEngine

  attr_reader :configuration, :simple_stats, :data_samples_count, :rule_combinations

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

    current_column = configuration.data_column_index
    header_string = Helper.cell_value data_sheet, configuration.data_row_index, current_column

    while !header_string.nil? do
      puts "Processing column #{header_string}"

      start_row_index = current_row = find_data_series_start data_sheet, configuration.data_row_index + 1, current_column
      cell_value = Helper.cell_value data_sheet, current_row, current_column

      while !cell_value.nil? do
        rules = configuration.rules[header_string]

        compute_simple_stats_for_cell cell_value, current_row, header_string, rules unless rules.nil?

        current_row = current_row + 1
        cell_value = Helper.cell_value data_sheet, current_row, current_column
      end

      current_column = current_column + 1
      header_string = Helper.cell_value data_sheet, configuration.data_row_index, current_column
    end

    @data_samples_count = current_row - start_row_index
    puts 'Simple stats computed'
  end

  def compute_bayes_correlations(workbook)
    @rule_combinations = build_rule_combinations.sort
    compute_probabilities @rule_combinations
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

  def build_rule_combinations()
    first_column = configuration.rules.keys.first
    other_columns = configuration.rules.keys - [first_column]

    all_combinations = []
    configuration.rules[first_column].each_pair do |first_label, rule|
      combinations = ["#{first_column}:#{first_label}"]
      other_columns.each do |column_name|
        new_combinations = []
        configuration.rules[column_name].keys.each do |current_label|
          combinations.each { |comb| new_combinations << comb + "|#{column_name}:#{current_label}" }
        end
        combinations.concat new_combinations
      end
      all_combinations.concat combinations.drop(1)
    end
    #Add single combination
    configuration.rules.each_pair do |column, rules|
      rules.each_pair { |label, rule| all_combinations << "#{column}:#{label}"}
    end
    all_combinations
  end

  def compute_probabilities(rule_combinations)

  end

end
