#!/usr/bin/ruby

require 'libarchive'
require 'xml'
require 'csv'

COL_SEP = ';'

if infile = ARGV[0]
	outfile = File.basename(infile,'.ods') + '.csv'
	entries = Archive.read_open_filename infile do |ar|
		while entry = ar.next_header
			if entry.pathname == 'content.xml'
				xml  = XML::Document.string ar.read_data
				CSV.open(outfile, "w", :col_sep => COL_SEP) do |csv|
					xml.find('/office:document-content/office:body/office:spreadsheet/table:table/table:table-row').each do |row|
						[row['number-rows-repeated'].to_i,1].max.times do
							csvrow = []
							row.each  do |cell|
								[cell['number-columns-repeated'].to_i,1].max.times do
									csvrow << cell.content
								end
								
							end
							csv << csvrow
						end
					end
				end
				break
			end
		end
	end
end
