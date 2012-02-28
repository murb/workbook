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
        puts " - Creating diff-table. Estimated time: #{aself.count*aself.first.count*0.0063}s "
        maxri = (aself.count-1)
        for ri in 0..maxri do
          row = diff_table[ri]
          row = diff_table[ri] = Workbook::Row.new(nil, diff_table)
          # aselfrow = aself[ri]
          #           aselfother
          srow = aself[ri]
          orow = aother[ri]
          
          iteration_cols.each_with_index do |ch, ci|
            scell = srow[ch]
            ocell = orow[ch]
            dcell = scell.nil? ? Workbook::Cell.new(nil) : scell
            if (scell == ocell)
              dcell.format = scell.format unless scell.nil?
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
      
      def compact_table
        self.delete_if{|r| r.nil? or r.compact.empty?}
        self
      end
      
      # aligns itself with another table, used by diff
      def align other, options={:sort=>true,:ignore_headers=>false}
        
        options = {:sort=>true,:ignore_headers=>false}.merge(options)
        puts " - Removing empty rows"
        
        iteration_cols = nil
        sother = other.clone.compact_table
        sself = self.clone.compact_table
        
        if options[:ignore_headers]
          sother.header = false
          sself.header = false
        end
        puts " - Sorting"
        
        sother = options[:sort] ? Workbook::Table.new(sother.sort) : sother
        sself = options[:sort] ? Workbook::Table.new(sself.sort) : sself
        
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
