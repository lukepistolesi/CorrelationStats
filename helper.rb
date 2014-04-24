class Helper

  def self.convert_letters_to_index(index_name)
    iteration_index = 0
    column_index = 0
    index_name.each_char do |char|
      column_index = column_index + (char.ord - 'A'.ord)*(26**iteration_index)
      iteration_index = iteration_index + 1
    end
    column_index
  end

  def self.empty_cell?(sheet, row, col)
    sheet.nil? || sheet[row].nil? || sheet[row][col].nil? || sheet[row][col].value.nil?
  end

  def self.cell_value(sheet, row, col)
    return nil if sheet.nil? || sheet[row].nil? || sheet[row][col].nil?
    sheet[row][col].value
  end

  def self.coordinates_to_sheet_and_cell(coordinates_with_sheet)
    coordinates_with_sheet.split ':'
  end

  def self.string_coordinates_to_indexes(string_coordinates)
    column = ''
    string_coordinates.each_char do |char|
      char =~ /[A-Z]/ ? column << char : break
    end

    row = string_coordinates[column.size .. -1]

    [Helper.convert_letters_to_index(column), Integer(row).to_i - 1, column, row]
  end
end