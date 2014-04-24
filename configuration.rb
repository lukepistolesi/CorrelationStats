# The following table is an example of what expected
# There should be at least a set of Rules and the table ends when
#   there is an empty line after the first rule section

# | Configuration         |

# | Correlation Threshold | 10.5                     |
# | Data Table Start      | Data:B3                  |
# | Categorized Columns   | Sex, Sick                |
# | Destination           | Sheet2:C3                |

# | Rules - Weight        |
# |                       | value < 70               | Light  |
# |                       | value > 70 && value < 80 | Medium |
# |                       | value > 80               | Heavy  |
# | Rules - Age           |
# |                       | value < 25               | Young  |
# |                       | value >= 25              | Old    |

require './helper.rb'
require 'sourcify'

class Configuration

  CONFIGURATION_TABLE_HEADER = 'Configuration'
  CONFIGURATION_TABLE_RULES_HEADER = 'Rules'

  PROPERTIES = %w(correlation_threshold data_table_start categorized_columns destination rules)

  PROPERTIES.each { |prop| attr_reader prop.to_sym }

  def initialize(spreadsheet, config_column, config_row)
    config_header = spreadsheet[config_column][config_row]
    if config_header.nil? || CONFIGURATION_TABLE_HEADER != config_header.value
      raise ArgumentError.new 'Configuration table not found'
    end

    last_row_index = load_settings spreadsheet, config_column, config_row
    load_rules spreadsheet, config_column, last_row_index
    validate
  end

  def to_s
    configuration_to_s = %Q{
| Configuration         |

| Correlation Threshold | #{correlation_threshold}
| Data Table Start      | #{data_table_start}
| Categorized Columns   | categorized_columns
| Destination           | destination
    }
    @rules.each_pair do |col_name, rules_hash|
      rules_to_s = %Q{\n| Rules - #{col_name.to_s} |}
      spaces = ' ' * (rules_to_s.size - 3)
      rule_spaces = rules_hash.values.collect { |proc| proc[:string].size }.max
      rules_hash.each_pair do |label, proc|
        rules_to_s.concat %Q{\n|#{spaces}| #{proc[:string]}#{' ' * (rule_spaces - proc[:string].size)} | #{label}}
      end
      configuration_to_s.concat rules_to_s
    end
    configuration_to_s
  end


  private

  PROPERTIES.each { |prop| attr_writer prop.to_sym }

  def load_settings(sheet, col, row)

    iterate_till_empty_rows(sheet, col, row) do |current_row_idx|

      cell_value = Helper.cell_value sheet, current_row_idx, col
      return false if is_rule_cell? cell_value

      attribute = cell_value.gsub(' ', '_').downcase
      value = Helper.cell_value sheet, current_row_idx, col + 1
      begin value = Float(value); rescue Exception; end

      self.send "#{attribute}=", value
    end

  end

  def load_rules(sheet, col, row)
    @rules = {}
    current_rules = nil
    current_rule_column_name = nil

    iterate_till_empty_rows(sheet, col + 1, row, false) do |current_row_idx|
      cell_value = Helper.cell_value sheet, current_row_idx, col
      if is_rule_cell? cell_value
        rule_column_name = cell_value.split('-')[1].strip
        current_rules = @rules[rule_column_name] = {}
        current_rule_column_name = rule_column_name
      elsif cell_value.nil?
        rule_string = Helper.cell_value sheet, current_row_idx, col + 1
        rule_label = Helper.cell_value sheet, current_row_idx, col + 2
        create_rule rule_label, rule_string, current_rules
      else
        raise ArgumentError.new "Unexpected Rule syntax at row #{current_row_idx + 1} and column #{col + 1}"
      end
    end
  end

  def validate()
    self.categorized_columns ||= []
    PROPERTIES.each do |prop|
      raise ArgumentError.new "Configuration Setting #{prop} not set" if self.send(prop.to_sym).nil?
    end
  end

  def is_rule_cell?(cell_value)
    !cell_value.nil? && cell_value.to_s.start_with?(CONFIGURATION_TABLE_RULES_HEADER)
  end

  def iterate_till_empty_rows(sheet, col, row, skip_empty_rows=true, &block)
    current_row = row + 1
    while true do
      current_cell_empty = Helper.empty_cell? sheet, current_row, col
      break if current_cell_empty && Helper.empty_cell?(sheet, current_row + 2, col)
      break if (!current_cell_empty || !skip_empty_rows) && !yield(current_row)
      current_row = current_row + 1
    end
    current_row
  end

  def create_rule(rule_label, rule_string, current_rules)
    current_rules[rule_label] = { proc: eval("Proc.new { |value| #{rule_string} }"), string: rule_string }
  end
end
