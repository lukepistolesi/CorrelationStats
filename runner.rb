require './command_line_options.rb'
require './correlation_calculator.rb'

opts = CommandLineOptions.new ARGV

puts "Ready to work on file #{opts.file_name} with configuration defined in sheet '#{opts.config_sheet}' - column idx '#{opts.config_column_idx}' - row idx '#{opts.config_row_idx}'"

calculator =
  CorrelationCalculator.new opts.file_name, opts.config_sheet, opts.config_row_idx, opts.config_column_idx

calculator.load_configuration

calculator.compute_conditional_probabilities

