require 'spec_helper'

describe StructuredReport::Report do
	it "should create an empty report" do
		report = StructuredReport::Report.new
		report.count.should == 0
	end

	it "should create an object with defiend columns" do
		report = StructuredReport::Report.new({
			:date => {:title => "Date"},
			:cost => {:title => "Cost",:type => :currency}
		})
		report.count.should == 0
		report.columns.count.should == 2
	end

	context 'having a report,' do
		before(:each) do
			@report = StructuredReport::Report.new({
				:date => {:title => "Date"},
				:cost => {:title => "Cost",:type => :currency}
			})
		end

		it 'can add a column' do
			@report.add_column(:quantity,{:title => "Quantity",:type => :numeric})
			@report.columns.count.should == 3
		end

		it 'can reference a column by id' do
			@report.columns[:cost][:type].should == :currency
			@report.columns[:cost][:data].should == []
		end

		it 'cannot add a duplicate column' do
			expect { @report.add_column(:date,{:title => "Date"}) }.to raise_exception(StructuredReport::Report::ColumnAlreadyExists)
		end

		it 'can add a row by hash' do
			@report.add_row({:date => "1/1/2001",:cost => "10"})
			@report.count.should == 1
			@report.columns[:date][:data][0].should == "1/1/2001"
			@report.columns[:cost][:data][0].should == "10"
		end

		it 'is not reliant on order for hashes' do
			@report.add_row({:cost => "10",:date => "1/1/2001"})
			@report.count.should == 1
			@report.columns[:date][:data][0].should == "1/1/2001"
			@report.columns[:cost][:data][0].should == "10"
		end

		it 'can add a row by array' do
			@report.add_row(["1/1/2001","10"])
			@report.count.should == 1
			@report.columns[:date][:data][0].should == "1/1/2001"
			@report.columns[:cost][:data][0].should == "10"
		end
	end

	context 'having a report with data' do
		before(:each) do
			@report = StructuredReport::Report.new({
				:date => {:title => "Date"},
				:cost => {:title => "Cost",:type => :currency}
			})

			@report.add_row({:date => "1/1/2001",:cost => "10"})
			@report.add_row({:date => "1/2/2001",:cost => "5"})
			@report.add_row({:date => "1/3/2001",:cost => "20"})
		end

		it 'can generate a CSV' do
			csv_string = @report.to_csv
			csv_string.should == "Date,Cost\n1/1/2001,$10.00\n1/2/2001,$5.00\n1/3/2001,$20.00\n"
		end

		it 'can generate an XML-based XLS' do
			xml_string = @report.to_xls
			xml_string.length.should > 0
		end

		it 'can generate a text-based table' do
			text_string = @report.to_text
			text_string.should == "Date     | Cost   \n------------------\n1/1/2001 | $10.00 \n1/2/2001 |  $5.00 \n1/3/2001 | $20.00 \n"
		end
	end
end