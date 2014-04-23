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
end