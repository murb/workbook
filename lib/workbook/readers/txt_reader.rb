# frozen_string_literal: true
# frozen_string_literal: true

module Workbook
  module Readers
    module TxtReader
      def load_txt text, options = {}
        csv = text
        parse_txt(csv, options)
      end

      def parse_txt(csv_raw, options = {})
        csv = csv_raw.split("\n").collect { |l| CSV.parse_line(l, col_sep: "\t") }
        self[0] = Workbook::Sheet.new(csv, self, parse_cells_on_batch_creation: true, cell_parse_options: {detect_date: true}) unless sheet.has_contents?
      end
    end
  end
end
