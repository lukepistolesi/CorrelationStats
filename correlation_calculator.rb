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

engine.compute_bayes_correlations workbook
rule_combos = engine.rule_combinations
puts "\n#{rule_combos.size} combinations built"
puts "\nConditional Probabilities"
#engine.conditionals.keys.sort.each do |combination|
engine.conditionals.sort_by { |comb, results| results[:probability] }.each do |key_val|
  #results = engine.conditionals[combination]
  results = key_val[1]
  posterior = results[:posterior]
  probability = results[:probability].round(2) * 100
  unless posterior.nil? || probability == 0.0
    posterior = posterior.gsub '|', '∩'
    puts "P(#{results[:prior]}|#{posterior}) = #{probability}%"
  end
end
