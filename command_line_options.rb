require './helper.rb'

class CommandLineOptions

  attr_reader :file_name, :config_sheet, :config_column, :config_row, :config_column_index, :config_row_index

  def initialize(command_line_arguments)
    if command_line_arguments.size != 2
      exception_msg =
        "\nPlease specify an excel file in xlsx format and the position of the configuration table" +
        "\nE.g. Spreadsheet.xlsx ConfigurationDataSheet1:F2"
      raise ArgumentError.new exception_msg
    end

    @file_name = command_line_arguments[0]

    splitted_config_coordinates = command_line_arguments[1].split ':'
    if splitted_config_coordinates.size != 2
      exception_msg =
        "\nThe coordinates for the configuration table are missing" +
        "\nE.g. ConfigurationDataSheet1:F2"
    end

    @config_sheet = splitted_config_coordinates[0]
    cell_coordinates = splitted_config_coordinates[1].upcase

    @config_column = ''
    cell_coordinates.each_char do |char|
      puts "Char is #{char}"
      char =~ /[A-Z]/ ? config_column << char : break
    end

    @config_row = cell_coordinates[config_column.size..-1]

    @config_column_index = Helper.convert_letters_to_index @config_column
    @config_row_index = Integer(@config_row).to_i - 1
  end

end