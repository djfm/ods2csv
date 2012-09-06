#!/usr/bin/ruby

require 'libarchive'
require 'xml'
require 'csv'

COL_SEP = ';'

num_rows = 0

if infile = ARGV[0]
	outfile = File.basename(infile,'.ods') + '.csv'
	entries = Archive.read_open_filename infile do |ar|
		while entry = ar.next_header
			if entry.pathname == 'content.xml'
				xml  = XML::Document.string ar.read_data
				CSV.open(outfile, "w", :col_sep => COL_SEP) do |csv|
					rows = xml.find('/office:document-content/office:body/office:spreadsheet/table:table/table:table-row')
					rows.each_with_index do |row, i|
						if i + 1 < rows.length then repeat_row = [row['number-rows-repeated'].to_i,1].max else repeat_row = 1 end
						repeat_row.times do
							csvrow = []
							row.each_with_index do |cell, j|
								if j + 1 < row.count then repeat_column = [cell['number-columns-repeated'].to_i,1].max else repeat_column = 1 end
								repeat_column.times do
									csvrow << cell.content
								end
							end
							csv << csvrow
							num_rows += 1
						end
					end
				end
				break
			end
		end
	end
	
	puts "Converted #{infile} to #{outfile}!"
	
else
	puts "Please provide a .ods file!"
end
