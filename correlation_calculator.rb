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

    puts "Configuration loaded"

    @engine = CorrelationEngine.new @configuration
  end

  def compute_conditional_probabilities()
    @engine.compute_simple_stats @workbook

    @engine.compute_bayes_correlations @workbook

    rule_combinations = @engine.rule_combinations
  end

  def save_results_in_workbook(options={print_zeroes: true})
    result_sheet = @workbook[@configuration.destination_sheet_name]
    if !options[:create_result_sheet] && result_sheet.nil?
      raise ArgumentError.new "Result spreadsheet does not exist: '#{@configuration.destination_sheet_name}'"
    elsif result_sheet.nil?
      result_sheet = @workbook.add_worksheet @configuration.destination_sheet_name
    end
    cur_column_name = ''
    cur_label = ''
    rule_label_counter = 0
    cur_row = @configuration.destination_row_idx
    cur_col = @configuration.destination_column_idx
    max_prob = 0.0
    max_prob_col_idx = nil
    max_prob_row_idx = nil

    engine.conditionals.sort_by { |comb, results| comb }.each do |key_val|
      first_rule, second_rule = Helper.split_combo_rule_in_head_and_tail key_val[0]
      rule_column, rule_label = first_rule.split ':'
      if rule_column != cur_column_name
        cur_column_name = rule_column
        cur_row = cur_row + 1
        rule_label_counter = 0
        Helper.set_cell_data result_sheet, cur_row, @configuration.destination_column_idx, rule_column
      end
      if rule_label != cur_label
        unless max_prob_col_idx.nil?
          Helper.set_cell_data result_sheet, max_prob_row_idx, max_prob_col_idx, max_prob, {fill_color: '0ba53d'}
        end
        rule_label_counter = rule_label_counter + 1
        cur_label = rule_label
        cur_row = cur_row + 1
        cur_col = @configuration.destination_column_idx + 2
        max_prob = 0.0
        max_prob_row_idx = max_prob_col_idx = nil
        Helper.set_cell_data result_sheet, cur_row, @configuration.destination_column_idx + 1, rule_label
      end
      if rule_label_counter == 1 && !second_rule.nil?
        Helper.set_cell_data result_sheet, cur_row - 1, cur_col, second_rule
      end
      value = key_val[1][:probability]
      if value != 0.0 || options[:print_zeroes]
        value = options[:percentage] ? (value * 100.0).round(2) : value.round(4)
        if max_prob < value && !second_rule.nil?
          max_prob = value
          max_prob_col_idx = cur_col
          max_prob_row_idx = cur_row
        end
      else
        value = ''
      end
      Helper.set_cell_data result_sheet, cur_row, cur_col, value
      cur_col = cur_col + 1
    end

    workbook.theme = RubyXL::Theme.defaults
    @workbook.write @configuration.destination_file
  end

  def to_s
    @configuration.to_s + "\n\n" + @engine.to_s
  end

  def conditional_probabilities; @engine.conditionals || {} end

end
