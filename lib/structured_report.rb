require 'csv'
require 'nokogiri'

module StructuredReport
	class Report
		class ColumnNotFound < StandardError; end
		class ColumnAlreadyExists < StandardError; end
		class MissingColumnData < StandardError; end

		attr_reader :count

		def initialize(columns = nil)
			@count = 0
			add_columns(columns)
		end

		def columns
			@columns ||= {}
		end

		def add_column(name,options = {})
			raise ColumnAlreadyExists if columns[name.to_sym]
			columns[name.to_sym] = StructuredReport::Column.new(options)
		end

		def add_columns(columns)
			if columns
				columns.each do |name,options|
					add_column(name,options)
				end
			end
		end

		def <<(data)
			add_row(data)
		end

		def add_row(data)
			raise MissingColumnData unless data.count >= columns.count

			if data.respond_to?(:keys)
				# Input is a hash
				i = 0
				data.each do |key,value|
					if columns[key.to_sym]
						columns[key.to_sym] << value
						i += 1
					end
				end

				raise MissingColumnData unless i == columns.count
			else
				i = 0
				columns.each do |id,column|
					column[:data][@count] = data[i]
					i += 1
				end
			end

			@count += 1
		end

		def each(&block)
			@count.times do |i|
				row = {}
				columns.each { |id,column| row[id] = column[:data][i] }
				block.call(row)
			end
		end

		def to_csv
			CSV.generate do |csv|
				headers = []
				columns.each do |id,column|
					headers << column[:title]
				end
				csv << headers

				self.each do |row|
					data = []
					row.each do |key,value|
						data << columns[key].format_value(value)
					end
					csv << data
				end
			end
		end

		def to_xls
			builder = Nokogiri::XML::Builder.new do |xml|
				xml.Workbook(
					'xmlns' => "urn:schemas-microsoft-com:office:spreadsheet",
  					'xmlns:o' => "urn:schemas-microsoft-com:office:office",
  					'xmlns:x' => "urn:schemas-microsoft-com:office:excel",
  					'xmlns:ss' => "urn:schemas-microsoft-com:office:spreadsheet",
  					'xmlns:html' => "http://www.w3.org/TR/REC-html40"
  				) {
					xml.Styles {
						xml.Style('ss:ID' => 'Header') {
							xml.Font('ss:Bold' => 1)
						}

						xml.Style('ss:ID' => 'Currency') {
							xml.NumberFormat('ss:Format' => "\"$\"#,##0.00")
						}
					}

					xml.Worksheet('ss:Name' => "Sheet 1") {
						xml.Table('ss:ExpandedColumnCount' => columns.count) {
							columns.each do |id,column|
								attributes = { 'ss:AutoFitWidth' => 1 }
								attributes['ss:Width'] = column[:width] if column[:width]
								attributes['ss:StyleID'] = 'Currency' if column[:type] == :currency
								xml.Column(attributes)
							end

							xml.Row {
								columns.each do |id,column|
									xml.Cell('ss:StyleID' => "Header") {
										xml.Data('ss:Type' => Column.xls_types[:string]) { 
											xml.text column[:title] 
										}
									}
								end								
							}

							self.each do |row|
								xml.Row {
									columns.each do |id,column|
										xml.Cell {
											xml.Data('ss:Type' => column[:xls_type]) { 
												xml.text column.format_value(row[id],:xls)
											}
										}
									end								
								}
							end
						}
					}
  				}
			end

			return builder.to_xml
		end

		def to_text

			widths = {}

			self.each do |row|
				c = 0
				row.each do |key,value|
					output = columns[key].format_value(value).to_s
					widths[c] = (widths[c] && widths[c] > output.length) ? widths[c] : output.length
					c += 1
				end
			end

			headers = ""
			divider = ""
			c = 0
			columns.each do |id,column|
				if c > 0 and c < widths.count
					headers += " | " 
					divider += "---"
				end

				widths[c] = (widths[c] && widths[c] > column[:title].length) ? widths[c] : column[:title].length
				headers += column[:title].ljust(widths[c])
				divider += "-"*widths[c]

				c += 1

				if c == widths.count
					headers += " " 
					divider += "-"
				end
			end

			output = "#{headers}\n#{divider}\n"

			self.each do |row|
				c = 0
				row.each do |key,value|
					output += " | " if c > 0 and c < widths.count

					column = columns[key]
					s = column.format_value(value)

					if [:numeric,:float,:currency].include?(column.type)
						output += s.rjust(widths[c])
					else
						output += s.ljust(widths[c])
					end

					c += 1

					output += " " if c == widths.count
				end

				output += "\n"
			end

			return output
		end
	end

	class Column
		class InvalidColumnType < StandardError; end

		attr_accessor :id,:title,:type,:width,:format
		attr_reader :data

		def initialize(options)
			options ||= {}

			options[:type] ||= :string
			raise InvalidColumnType unless Column.column_types.include?(options[:type])

			@title = (options[:title] || "")
			@type = options[:type]
			@format = options[:format]
			@width = options[:width]
			@data = []
		end

		def format_value(value,type = :csv,output_key = :text)
			if value.respond_to?(:keys)
				display_value = value[output_key]
			else
				display_value = value
			end

			return "" unless display_value

			begin
				output = sprintf(format,display_value)
				output.gsub!(/[^0-9.-]/,"") if type == :xls and xls_type == 'Number'
			rescue => e
				raise "Error Rendering Value: #{value}, #{display_value}, #{type}, #{output_key}, #{format} - #{e}"
			end

			return output
		end

		def format
			(@format || Column.formats[self.type])
		end

		def [](attribute)
			self.send(attribute.to_s)
		end

		def <<(value)
			@data << value
		end

		def xls_type
			Column.xls_types[self.type]
		end

		def self.column_types
			[:string,:numeric,:float,:currency]
		end

		def self.xls_types
			{
				:string => 'String',
				:numeric => 'Number',
				:float => 'Number',
				:currency => 'Number'
			}
		end

		def self.formats
			{
				:string => '%s',
				:numeric => '%d',
				:float => '%f',
				:currency => '$%.2f'
			}
		end
	end
end