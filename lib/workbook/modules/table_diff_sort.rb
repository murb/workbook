# Adds diffing and sorting functions
module Workbook
	module Modules
		module TableDiffSort
      # create an overview of the differences between itself with another table, returns a book with a single sheet and table (containing the diffs)
      def diff other, options={:sort=>true,:ignore_headers=>false}
        puts "#diff"
        
        aligned = align(other, options)
        aself = aligned[:self]
        aother = aligned[:other]
        
        
        iteration_cols = []
        if options[:ignore_headers]
          iteration_cols = [aother.first.count,aself.first.count].max.times.collect
        else
          iteration_cols = (aother.header.to_symbols+aother.header.to_symbols).uniq
        end
        diff_table = diff_template.sheet.table
        iteration_rows = aself.count.times
        puts " - Creating diff-table. Estimated time: #{aself.count*aself.first.count*0.0063}s "
        
        iteration_rows.each do |ri|
          row = diff_table[ri]
          row = diff_table[ri] = Workbook::Row.new(nil, diff_table)
          if ri == 0 and !options[:ignore_headers]
            row.format = diff_template.template.create_or_find_format_by 'header'
          end
          iteration_cols.each_with_index do |ch, ci|
            scell = aself[ri][ch]
            ocell = aother[ri][ch]
            if scell == ocell or (scell and ocell and scell.value == ocell.value)
               if scell
                 dcell = scell
                 dcell.format = scell.format
               else
                 dcell = Workbook::Cell.new(nil)
               end
            elsif scell.nil? or scell.value.nil?
              dcell = Workbook::Cell.new "(was: #{ocell.to_s})"
              dcell.format = diff_template.template.create_or_find_format_by 'destroyed'
            elsif ocell.nil? or ocell.value.nil?
              dcell = scell.clone
              f = diff_template.template.create_or_find_format_by 'created', scell.format[:number_format]
              f[:number_format] = scell.format[:number_format]
              dcell.format = f
            elsif scell.value != ocell.value
              dcell = Workbook::Cell.new "#{scell.to_s} (was: #{ocell.to_sl})"
              f = diff_template.template.create_or_find_format_by 'updated'
              dcell.format = f
            end
            
            row[ci]=dcell
          end
        end

        diff_template
      end
      
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
        @diff_template = diffbook
        return diffbook
      end
      
      # aligns itself with another table, used by diff
      def align other, options={:sort=>true,:ignore_headers=>false}
        puts " - Sorting"
        options = {:sort=>true,:ignore_headers=>false}.merge(options)
        
        iteration_cols = nil
        sother = other.clone
        sself = self.clone
        if options[:ignore_headers]
          sother.header = false
          sself.header = false
        end
        
        sother = options[:sort] ? sother.sort : sother
        sself = options[:sort] ? sself.sort : sself
        
        iteration_rows =  [sother.count,sself.count].max.times.collect
        puts " - Aligning"

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
