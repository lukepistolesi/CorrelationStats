require 'rubyXL'
require 'ruby2ruby'
require './numeric.rb'
require './command_line_options.rb'
require './configuration.rb'
require './correlation_engine.rb'

cmd_line_opts = CommandLineOptions.new ARGV

puts "Ready to work on file #{cmd_line_opts.file_name} with configuration defined in sheet '#{cmd_line_opts.config_sheet}' - column index '#{cmd_line_opts.config_column_index}' - row index '#{cmd_line_opts.config_row_index}'"
# puts "  with configuration stored in #{cmd_line_opts.config_sheet}"
# puts "  column #{cmd_line_opts.config_column} -- #{cmd_line_opts.config_column_index}"
# puts "  row #{cmd_line_opts.config_row} -- #{cmd_line_opts.config_row_index}"
workbook = RubyXL::Parser.parse cmd_line_opts.file_name
#puts "Workbook loaded with #{workbook.worksheets.size} spreadsheets"

config_sheet = workbook[cmd_line_opts.config_sheet]
configuration = Configuration.new config_sheet, cmd_line_opts.config_column_index, cmd_line_opts.config_row_index

puts "Configuration loaded: #{configuration}"

engine = CorrelationEngine.new configuration
engine.compute_simple_stats workbook

puts engine.to_s

#notnil = sheet.sheet_data.select { |row| !row.nil? }.size
# rowno = 0
# (0..sheet.sheet_data.size).each do |idx|
#   row = sheet.sheet_data[idx]
#   rowno = rowno + 1 unless row.nil?
# end
# puts "Number of rows is #{rowno}"

# (0..sheet.sheet_data.size).each do |idx|
#   next if (row = sheet.sheet_data[idx]).nil?
#   puts "Row #{idx}\n  has size #{row.size}\n  Second column cell is #{row[1].value unless row[1].nil?}"
# end
