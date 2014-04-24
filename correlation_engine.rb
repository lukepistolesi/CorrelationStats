class CorrelationEngine

  attr_reader :configuration

  def initialize(configuration)
    @configuration = configuration
  end

  def compute_simple_stats(workbook)
    data_sheet = workbook[configuration.data_sheet_name]
    puts "Data sheet loaded for '#{configuration.data_sheet_name}': #{data_sheet}"
    puts 'Starting computing simple stats'

    puts 'Simple stats computed'
  end
end