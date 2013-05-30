require 'spreadsheet'
require 'csv'


class ExcelMerge
  TEMPLATE_HEADER_ROW = 8
  TEMPLATE_OFFSET     = 9  # Point at which we insert into the template

  class << self
    def merge_template_with_csv(template, csv)
      template    = Spreadsheet.open(template)
      worksheet   = template.worksheets.first
      xls_headers = worksheet.row(TEMPLATE_HEADER_ROW).to_a.map {|c| c.downcase unless c.nil? }
      csv         = CSV.parse(csv)
      csv_headers = csv.shift.map(&:downcase)

      csv.each_with_index do |csv_row, row_index|
        csv_row.each_with_index do |csv_cell, col_index|
          next if csv_cell.nil?

          xls_col = xls_headers.find_index(csv_headers[col_index]) or
            raise 'CSV column not found in template.'

          worksheet[row_index + TEMPLATE_OFFSET, xls_col] = csv_cell
        end
      end

      template
    end
  end
end
