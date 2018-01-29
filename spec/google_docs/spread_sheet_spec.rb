require 'spec_helper'

describe GoogleDocs::SpreadSheet do
  let!(:service)             { double :service, get_spreadsheet: spreadsheet }
  let!(:google_access_token) { double :google_access_token }
  let!(:sheet1)              { double :sheet1, properties: double(:sheet1_properties, sheet_id: 1) }
  let!(:sheet2)              { double :sheet2, properties: double(:sheet2_properties, sheet_id: 2) }
  let!(:sheet)               { double :sheet }
  let!(:spreadsheet) do
    double :spreadsheet, spreadsheet_id: 'WE123', sheets: [sheet1, sheet2],
                         properties: double(:spreadsheet_properties, title: 'Sheet 1')
  end

  before do
    allow(Google::Apis::SheetsV4::SheetsService).to receive(:new).and_return(service)
    allow(service).to receive(:authorization=).and_return(google_access_token)
  end

  describe '#sheets' do
    before do
      allow(GoogleDocs::Sheet).to receive(:new).and_return(sheet)
      GoogleDocs::SpreadSheet.new('google_access_token', 'WE123').sheets
    end

    it { expect(GoogleDocs::Sheet).to have_received(:new).with(spreadsheet_id: 'WE123', sheet_id: 1, service: service) }
    it { expect(GoogleDocs::Sheet).to have_received(:new).with(spreadsheet_id: 'WE123', sheet_id: 2, service: service) }
  end

  describe '#properties' do
    before { GoogleDocs::SpreadSheet.new('google_access_token', 'WE123').properties }

    it { expect(spreadsheet).to have_received(:properties) }
  end
end
