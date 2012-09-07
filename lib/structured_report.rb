require 'csv'

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
			options ||= {}

			raise ColumnAlreadyExists if columns[name.to_sym]
			columns[name.to_sym] = {
				:title => (options[:title] || ""),
				:type => (options[:type] || :string),
				:format => (options[:format] || "%s"),
				:data => []
			}
		end

		def add_columns(columns)
			if columns
				columns.each do |name,options|
					add_column(name,options)
				end
			end
		end

		def add_row(data)
			raise MissingColumnData unless data.count == columns.count

			if data.respond_to?(:keys)
				# Input is a hash

				data.each do |key,value|
					raise ColumnNotFound unless columns[key.to_sym]
					columns[key.to_sym][:data][@count] = value
				end
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
						data << sprintf(columns[key][:format],value)
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
							xml.NumberFormat('ss:Format' => "&quot;$&quot;#,##0.00")
						}
					}

					xml.Worksheet('ss:Name' => "Sheet 1") {
						xml.Table('ss:ExpandedColumnCount' => columns.count) {
							columns.each do |id,column|
								attributes = { 'ss:AutoFitWidth' => 1 }
								attributes['ss:Width'] = column[:width] if column[:width]
								attributes['ss:Type'] = 'Currency' if column[:type] == :currency
								xml.Column(attributes)
							end

							xml.Row {
								columns.each do |id,column|
									xml.Cell('ss:StyleID' => "Header") {
										xml.Data('ss:Type' => Report.xlsTypes[:string]) { 
											xml.text column[:title] 
										}
									}
								end								
							}

							self.each do |row|
								xml.Row {
									columns.each do |id,column|
										xml.Cell {
											xml.Data('ss:Type' => Report.xlsTypes[column[:type]]) { 
												xml.text sprintf(column[:format],row[id])
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

		def self.xlsTypes
			{
				:string => 'String',
				:numeric => 'Numeric',
				:currency => 'Numeric'
			}
		end
	end
end