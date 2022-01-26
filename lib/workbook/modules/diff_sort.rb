# frozen_string_literal: true
# frozen_string_literal: true

module Workbook
  module Modules
    # Adds essential diffing and comparing support, as well as diffing entire books
    module BookDiffSort
      module ClassMethods
        # Return template table to write the diff result in; in case non exists a default is generated.
        #
        # @return [Workbook::Table] the empty table, linked to a book
        def new_diff_template
          diffbook = Workbook::Book.new
          template = diffbook.template
          f = template.create_or_find_format_by "destroyed"
          f[:background_color] = :red
          f = template.create_or_find_format_by "updated"
          f[:background_color] = :yellow
          f = template.create_or_find_format_by "created"
          f[:background_color] = :lime
          f = template.create_or_find_format_by "header"
          f[:rotation] = 72
          f[:font_weight] = :bold
          f[:height] = 80
          diffbook
        end
      end

      def self.included(base)
        base.extend(ClassMethods)
      end

      # Diff an entire workbook against another, sheet by sheet
      #
      # @param [Workbook::Book] to_workbook to compare against
      # @return [Workbook::Book] workbook with compared result
      def diff to_workbook, options = {sort: true, ignore_headers: false}
        diff_template = Workbook::Book.new_diff_template
        each_with_index do |from_sheet, sheet_index|
          to_sheet = to_workbook[sheet_index]
          if to_sheet
            from_table = from_sheet.table
            to_table = to_sheet.table
            diff_table_template = diff_template.create_or_open_sheet_at(sheet_index).table
            from_table.diff_template = diff_table_template
            from_table.diff(to_table, options)
          end
        end
        diff_template # the template has been filled in the meanwhile, not to use as a template anymore... :)
      end
    end

    # Adds diffing and sorting functions
    module TableDiffSort
      # create an overview of the differences between itself with another 'previous' table, returns a book with a single sheet and table (containing the diffs)
      #
      # @return [Workbook::Table] the return result
      def diff other, options = {}
        options = {sort: true, ignore_headers: false}.merge(options)

        aligned = align(other, options)
        aself = aligned[:self]
        aother = aligned[:other]

        iteration_cols = []
        iteration_cols = if options[:ignore_headers]
          [aother.first.count, aself.first.count].max.times.collect
        else
          (aother.header.to_symbols + aother.header.to_symbols).uniq
        end

        diff_table = diff_template
        maxri = (aself.count - 1)

        (0..maxri).each do |ri|
          row = diff_table.rows[ri] = Workbook::Row.new(nil, diff_table)
          srow = aself.rows[ri]
          orow = aother.rows[ri]

          iteration_cols.each_with_index do |ch, ci|
            scell = srow[ch]
            ocell = orow[ch]
            row[ci] = create_diff_cell(scell, ocell)
          end
        end
        unless options[:ignore_headers]
          diff_table[0].format = diff_template.template.create_or_find_format_by "header"
        end

        diff_table
      end

      # Return template table to write the diff result in; in case non exists a default is generated.
      #
      # @return [Workbook::Table] the empty table, linked to a book
      def diff_template
        return @diff_template if defined?(@diff_template)
        diffbook = Workbook::Book.new_diff_template
        difftable = diffbook.sheet.table
        @diff_template ||= difftable
      end

      # Set the template table to write the diff result in; in case non exists a default is generated. Make sure that
      # the following formats exists: destroyed, updated, created and header.
      #
      # @param [Workbook::Table] table to diff inside
      # @return [Workbook::Table] the passed table
      def diff_template= table
        @diff_template = table
      end

      # aligns itself with another table, used by diff
      #
      # @param [Workbook::Table] other table to align with
      # @param [Hash] options default to: `{:sort=>true,:ignore_headers=>false}`
      def align other, options = {}
        options = {sort: true, ignore_headers: false}.merge(options)

        sother = other.clone.remove_empty_lines!
        sself = clone.remove_empty_lines!

        if options[:ignore_headers]
          sother.header = false
          sself.header = false
        end

        sother = options[:sort] ? sother.sort : sother
        sself = options[:sort] ? sself.sort : sself

        row_index = 0
        while (row_index < [sother.count, sself.count].max) && (row_index < other.count + count)
          row_index = align_row(sself, sother, row_index)
        end

        {self: sself, other: sother}
      end

      def sort
        clone.sort!
      end

      def sort!
        header_row = @rows.delete_at(header_row_index) if header
        @rows = [header_row] + @rows.sort
        self
      end

      private

      # for use in the align 'while' loop
      def align_row sself, sother, row_index
        asd = 0
        if sself.rows[row_index] && sother.rows[row_index]
          asd = sself.rows[row_index].key <=> sother.rows[row_index].key
        elsif sself.rows[row_index]
          asd = -1
        elsif sother.rows[row_index]
          asd = 1
        end
        if (asd == -1) && insert_placeholder?(sother, sself, row_index)
          sother.rows.insert row_index, placeholder_row
          row_index -= 2
        elsif (asd == 1) && insert_placeholder?(sother, sself, row_index)
          sself.rows.insert row_index, placeholder_row
          row_index -= 2
        end

        row_index + 1
      end

      def insert_placeholder? sother, sself, row_index
        (sother.rows[row_index].nil? || !sother.rows[row_index].placeholder?) &&
          (sself.rows[row_index].nil? || !sself.rows[row_index].placeholder?)
      end

      # returns a placeholder row, for internal use only
      def placeholder_row
        return @placeholder_row if defined?(@placeholder_row) && !@placeholder_row.nil?

        @placeholder_row = Workbook::Row.new [nil]
        @placeholder_row.placeholder = true
        @placeholder_row
      end

      # creates a new cell describing the difference between two cells
      #
      # @return [Workbook::Cell] the diff cell
      def create_diff_cell(scell, ocell)
        dcell = scell.nil? ? Workbook::Cell.new(nil) : scell
        if scell == ocell
          dcell.format = scell.format if scell
        elsif scell.nil?
          dcell = Workbook::Cell.new "(was: #{ocell})"
          dcell.format = diff_template.template.create_or_find_format_by "destroyed"
        elsif ocell.nil?
          dcell = scell.clone
          fmt = scell.nil? ? :default : scell.format[:number_format]
          f = diff_template.template.create_or_find_format_by "created", fmt
          f[:number_format] = scell.format[:number_format]
          dcell.format = f
        elsif scell != ocell
          dcell = Workbook::Cell.new "#{scell} (was: #{ocell})"
          f = diff_template.template.create_or_find_format_by "updated"
          dcell.format = f
        end

        dcell
      end
    end
  end
end
