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

class Configuration

  CONFIGURATION_TABLE_HEADER = 'Configuration'
  CONFIGURATION_TABLE_RULES_HEADER = 'Rules'

  PROPERTIES = %w(correlation_threshold data_table_start categorized_columns destination)

  PROPERTIES.each { |prop| attr_reader prop.to_sym }

  def initialize(spreadsheet, config_column, config_row)
    config_header = spreadsheet[config_column][config_row]
    if config_header.nil? || CONFIGURATION_TABLE_HEADER != config_header.value
      raise ArgumentError.new 'Configuration table not found'
    end

    load_settings(spreadsheet, config_column, config_row)
    validate
  end


  private

  PROPERTIES.each { |prop| attr_writer prop.to_sym }

  def load_settings(sheet, col, row)
    # if Helper.empty_cell?(sheet, row + 1, col) || Helper.empty_cell?(sheet, row + 2, col)
    #   raise ArgumentError.new 'Configuration table is empty'
    # end

    current_row = row + 1
    while true do
      current_cell_empty = Helper.empty_cell? sheet, current_row, col
      break if (current_cell_empty && Helper.empty_cell?(sheet, current_row + 2, col)) ||
                is_rule_cell?(Helper.cell_value(sheet, current_row, col))

      unless current_cell_empty
        attribute = Helper.cell_value(sheet, current_row, col).gsub(' ', '_').downcase
        value = Helper.cell_value sheet, current_row, col + 1
        puts "Attribute #{attribute} -- #{value}"
        begin value = Float(value); rescue Exception; end

        self.send "#{attribute}=", value
      end

      current_row = current_row + 1
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
end
