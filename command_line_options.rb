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

    @config_sheet, cell_coordinates = Helper.coordinates_to_sheet_and_cell command_line_arguments[1]

    if @config_sheet.nil? || cell_coordinates.nil?
      exception_msg =
        "\nThe coordinates for the configuration table are missing" +
        "\nE.g. ConfigurationDataSheet1:F2"
    end

    @config_column_index, @config_row_index, @config_column, @config_row = Helper.string_coordinates_to_indexes cell_coordinates
  end

end