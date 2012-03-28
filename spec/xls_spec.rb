require 'spec_helper'

describe ToSpreadsheet::XLS do
  let(:html) { Haml::Engine.new(TEST_HAML).render }
  let(:workbook) { ToSpreadsheet::XLS.to_io(html) }
  let(:spreadsheet) {
    tmp_file = Tempfile.new(%w[to_spreadsheet .xlsx])
    IO.copy_stream(workbook.binmode, tmp_file.binmode)
    tmp_file.close
    RubyXL::Parser.parse(tmp_file.path)
  }

  it 'creates multiple worksheets' do
    spreadsheet.should have(2).worksheets
  end

  it 'supports num format' do
    spreadsheet.worksheets[0][1][1].value.should == 20
  end

  it 'support float format' do
    spreadsheet.worksheets[1][1][1].value.should be_a(Float)
  end

  it 'supports date format' do
    Date.parse(spreadsheet.worksheets[0][1][2].value).should be_a(Date)
  end
end

TEST_HAML = <<-HAML

%table
  %caption A worksheet
  %thead
    %tr
      %th Name
      %th Age
      %th Date
  %tbody
    %tr
      %td Gleb
      %td.num 20
      %td.date 27/05/1991
    %tr
      %td John
      %td.num 21
      %td.date 01/05/1990

%table
  %caption Another worksheet
  %thead
    %tr
      %th Name
      %th Age
      %th Date
  %tbody
    %tr
      %td Alice
      %td.float 19.5
      %td.date 10/05/1991

HAML
