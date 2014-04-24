class CorrelationEngine

  attr_reader :configuration

  def initialize(configuration)
    @configuration = configuration
  end

  def compute_simple_stats(workbook)
    data_sheet = workbook[configuration.data_sheet_name]
    puts "Data sheet loaded for '#{configuration.data_sheet_name}': #{data_sheet}"
    puts 'Starting computing simple stats'

    current_column = configuration.data_column_index
    header_string = Helper.cell_value data_sheet, configuration.data_row_index, current_column

    while !header_string.nil? do
      puts "Processing column #{header_string}"

      current_row = find_data_series_start data_sheet, configuration.data_row_index + 1, current_column
      cell_value = Helper.cell_value data_sheet, current_row, current_column

      while !cell_value.nil? do
        compute_simple_stats_for_cell cell_value, header_string, configuration.rules

        current_row = current_row + 1
        cell_value = Helper.cell_value data_sheet, current_row, current_column
      end

      current_column = current_column + 1
      header_string = Helper.cell_value data_sheet, configuration.data_row_index, current_column
    end

    puts 'Simple stats computed'
  end

  private

  def find_data_series_start(data_sheet, row, col)
    begin
      cell_value = Helper.cell_value data_sheet, row, col
    end while cell_value.nil? && (row = row + 1)
    row
  end

  def compute_simple_stats_for_cell(cell_value, header_string, rules_collection)
    rules = rules_collection[header_string]
    if rules.nil?
      puts "  WARNING: No rules defined for column #{header_string}"
    else
      puts "  Stats for #{cell_value} with #{rules.size unless rules.nil?} rules"
    end
  end
end