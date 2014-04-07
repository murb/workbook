# -*- encoding : utf-8 -*-
module Workbook
  module Modules
    # Adds diffing and sorting functions
    module TableDiffSort
      # create an overview of the differences between itself with another 'previous' table, returns a book with a single sheet and table (containing the diffs)
      #
      # @return [Workbook::Table] the return result
      def diff other, options={:sort=>true,:ignore_headers=>false}

        aligned = align(other, options)
        aself = aligned[:self]
        aother = aligned[:other]
        iteration_cols = []
        if options[:ignore_headers]
          iteration_cols = [aother.first.count,aself.first.count].max.times.collect
        else
          iteration_cols = (aother.header.to_symbols+aother.header.to_symbols).uniq
        end
        diff_table = diff_template
        maxri = (aself.count-1)
        for ri in 0..maxri do
          row = diff_table[ri]
          row = diff_table[ri] = Workbook::Row.new(nil, diff_table)
          srow = aself[ri]
          orow = aother[ri]

          iteration_cols.each_with_index do |ch, ci|
            scell = srow[ch]
            ocell = orow[ch]
            dcell = scell.nil? ? Workbook::Cell.new(nil) : scell
            if (scell == ocell)
              dcell.format = scell.format if scell
            elsif scell.nil?
              dcell = Workbook::Cell.new "(was: #{ocell.to_s})"
              dcell.format = diff_template.template.create_or_find_format_by 'destroyed'
            elsif ocell.nil?
              dcell = scell.clone
              fmt = scell.nil? ? :default : scell.format[:number_format]
              f = diff_template.template.create_or_find_format_by 'created', fmt
              f[:number_format] = scell.format[:number_format]
              dcell.format = f
            elsif scell != ocell
              dcell = Workbook::Cell.new "#{scell.to_s} (was: #{ocell.to_s})"
              f = diff_template.template.create_or_find_format_by 'updated'
              dcell.format = f
            end

            row[ci]=dcell
          end
        end
        if !options[:ignore_headers]
          diff_table[0].format = diff_template.template.create_or_find_format_by 'header'
        end

        diff_table
      end

      # Return template table to write the diff result in; in case non exists a default is generated.
      #
      # @return [Workbook::Table] the empty table, linked to a book
      def diff_template
        return @diff_template if @diff_template
        diffbook = Workbook::Book.new
        difftable = diffbook.sheet.table
        template = diffbook.template
        f = template.create_or_find_format_by 'destroyed'
        f[:background_color]=:red
        f = template.create_or_find_format_by 'updated'
        f[:background_color]=:yellow
        f = template.create_or_find_format_by 'created'
        f[:background_color]=:lime
        f = template.create_or_find_format_by 'header'
        f[:rotation] = 72
        f[:font_weight] = :bold
        f[:height] = 80
        @diff_template ||= difftable
      end

      # Set the template table to write the diff result in; in case non exists a default is generated. Make sure that
      # the following formats exists: destroyed, updated, created and header.
      #
      # @param [Workbook::Table] table to diff inside
      # @return [Workbook::Table] the passed table
      def diff_template= table
        @diff_template= table
      end

      # aligns itself with another table, used by diff
      #
      # @param [Workbook::Table] other table to align with
      # @param [Hash] options default to: `{:sort=>true,:ignore_headers=>false}`
      def align other, options={:sort=>true,:ignore_headers=>false}

        options = {:sort=>true,:ignore_headers=>false}.merge(options)

        sother = other.clone.remove_empty_lines!
        sself = self.clone.remove_empty_lines!

        if options[:ignore_headers]
          sother.header = false
          sself.header = false
        end

        sother = options[:sort] ? Workbook::Table.new(sother.sort) : sother
        sself = options[:sort] ? Workbook::Table.new(sself.sort) : sself

        row_index = 0
        while row_index < [sother.count,sself.count].max and row_index < other.count+self.count do
          row_index = align_row(sself, sother, row_index)
        end

        {:self=>sself, :other=>sother}
      end

      # for use in the align 'while' loop
      def align_row sself, sother, row_index
        asd = 0
        if sself[row_index] and sother[row_index]
          asd = sself[row_index].key <=> sother[row_index].key
        elsif sself[row_index]
          asd = -1
        elsif sother[row_index]
          asd = 1
        end
        if asd == -1 and insert_placeholder?(sother, sself, row_index)
          sother.insert row_index, placeholder_row
          row_index -=1
        elsif asd == 1 and insert_placeholder?(sother, sself, row_index)
          sself.insert row_index, placeholder_row
          row_index -=1
        end

        row_index += 1
      end

      def insert_placeholder? sother, sself, row_index
        (sother[row_index].nil? or !sother[row_index].placeholder?) and
        (sself[row_index].nil? or !sself[row_index].placeholder?)
      end

      # returns a placeholder row, for internal use only
      def placeholder_row
        if @placeholder_row != nil
          return @placeholder_row
        else
          @placeholder_row = Workbook::Row.new [nil]
          placeholder_row.placeholder = true
          return @placeholder_row
        end
      end
    end
  end
end
