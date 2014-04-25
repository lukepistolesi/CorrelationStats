require 'rubyXL'
require 'ruby2ruby'
require './numeric.rb'
require './command_line_options.rb'
require './configuration.rb'
require './correlation_engine.rb'

class CorrelationCalculator

  attr_reader :workbook, :config_sheet, :config_row, :config_col, :engine, :rule_combinations

  def initialize(excel_spreadsheet_file, config_sheet_name, config_row_idx, config_column_idx)
    @workbook = RubyXL::Parser.parse excel_spreadsheet_file
    @config_sheet = @workbook[config_sheet_name]
    @config_row = config_row_idx
    @config_col = config_column_idx
  end

  def load_configuration()
    @configuration = Configuration.new @config_sheet, @config_col, @config_row

    puts "Configuration loaded: #{@configuration}"

    @engine = CorrelationEngine.new @configuration
  end

  def compute_conditional_probabilities()
    @engine.compute_simple_stats @workbook
    puts @engine.to_s

    @engine.compute_bayes_correlations @workbook

    rule_combinations = @engine.rule_combinations
    puts "\n#{rule_combinations.size} combinations built"

    puts "\nConditional Probabilities"
    @engine.conditionals.sort_by { |comb, results| results[:probability] }.each do |key_val|
      results = key_val[1]
      posterior = results[:posterior]
      probability = results[:probability].round(2) * 100
      unless posterior.nil? || probability == 0.0
        posterior = posterior.gsub '|', '∩'
        puts "P(#{results[:prior]}|#{posterior}) = #{probability}%"
      end
    end
  end

  def to_s
    @configuration.to_s + "\n\n" + @engine.to_s
  end

  def conditional_probabilities; @engine.conditionals || {} end

end
