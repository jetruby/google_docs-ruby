require 'spec_helper'

describe GoogleDocs::Sheet do
  let(:service) do
    double :service, get_spreadsheet: spreadsheet, get_spreadsheet_values: double(:spreadsheet_values, values: results)
  end
  let(:sheet1) do
    double :sheet1, properties: double(:sheet1_properties, sheet_id: 1, title: 'First sheet name',
                                                           grid_properties: double(:grid_properties, column_count: 5, row_count: 3))
  end
  let(:spreadsheet) do
    double :spreadsheet, spreadsheet_id: 'WE123', sheets: [sheet1],
                         properties: double(:spreadsheet_properties, title: 'Sheet 1')
  end
  let(:results) do
    [
      %w[Date Developer Task Notes Hours],
      ['09/14/2016', 'Author 1', 'Testing Google API', 'API is ready', '1.15'],
      ['Total:', '', '', '', '1.15']
    ]
  end
  let(:request) { double :request }

  subject { described_class.new(spreadsheet_id: 'WE123', sheet_id: '1') }

  before do
    allow(Google::Apis::SheetsV4::SheetsService).to receive(:new).and_return(service)
    allow(Google::Apis::SheetsV4::Request).to receive(:new).and_return(request)
    allow(service).to receive(:authorization=).and_return('google_access_token')
  end

  describe '#get_rows_values' do
    it { expect(subject.get_rows_values).to eq results }
  end

  describe '#new' do
    it do
      expect{described_class.new(spreadsheet_id: 'WE123', sheet_id: 'Invalid id')}.to raise_error(ArgumentError)
    end
  end

  describe '#download_pdf' do
    let(:google_url)  { 'https://docs.google.com/spreadsheets/d/WE123/export?exportFormat=pdf&gid=1' }
    let(:google_data) { double :google_data }

    before { allow(service).to receive(:http).with(:get, google_url) { google_data } }

    specify do
      subject.download_pdf { |file| file }

      expect(service).to have_received(:http).with(:get, google_url) { google_data }
    end
  end

  describe '#setup_columns_width' do
    let(:data) { { 'A' => 40, 'B' => 60, 'C' => 150 } }
    let(:dimension_properties) { double :dimension_properties }
    let(:dimension_range) { double :dimension_range }
    let(:updated_dimension_properties_request) { double :updated_dimension_properties_request }

    before do
      allow(Google::Apis::SheetsV4::UpdateDimensionPropertiesRequest).to receive(:new)
        .and_return(updated_dimension_properties_request)
      allow(Google::Apis::SheetsV4::DimensionRange).to receive(:new).and_return(dimension_range)
      allow(Google::Apis::SheetsV4::DimensionProperties).to receive(:new).and_return(dimension_properties)
    end

    it { expect(subject.setup_columns_width(data)).to eq true }

    specify do
      subject.setup_columns_width(data)

      expect(Google::Apis::SheetsV4::DimensionProperties).to have_received(:new).with(pixel_size: 40)
      expect(Google::Apis::SheetsV4::DimensionProperties).to have_received(:new).with(pixel_size: 60)
      expect(Google::Apis::SheetsV4::DimensionProperties).to have_received(:new).with(pixel_size: 150)
      expect(Google::Apis::SheetsV4::DimensionRange).to have_received(:new)
        .with(sheet_id: 1, dimension: 'COLUMNS', start_index: 0, end_index: 1)
      expect(Google::Apis::SheetsV4::DimensionRange).to have_received(:new)
        .with(sheet_id: 1, dimension: 'COLUMNS', start_index: 1, end_index: 2)
      expect(Google::Apis::SheetsV4::DimensionRange).to have_received(:new)
        .with(sheet_id: 1, dimension: 'COLUMNS', start_index: 2, end_index: 3)
      expect(Google::Apis::SheetsV4::UpdateDimensionPropertiesRequest).to have_received(:new).exactly(3).times
    end
  end

  describe '#update_grid_properties' do
    let(:data) do
      {
        row_count:           3,
        column_count:        3,
        frozen_row_count:    1,
        frozen_column_count: 0,
        hide_gridlines:      false
      }
    end
    let(:updated_sheet_properties_request) { double :updated_sheet_properties_request }
    let(:sheet_properties) { double :sheet_properties }
    let(:grid_properties)  { double :grid_properties }
    let(:sheet_property_fields) do
      'gridProperties.RowCount,gridProperties.ColumnCount,gridProperties.FrozenRowCount,' \
      'gridProperties.FrozenColumnCount,gridProperties.HideGridlines'
    end

    before do
      allow(Google::Apis::SheetsV4::UpdateSheetPropertiesRequest).to receive(:new) { updated_sheet_properties_request }
      allow(Google::Apis::SheetsV4::SheetProperties).to receive(:new)              { sheet_properties }
      allow(Google::Apis::SheetsV4::GridProperties).to receive(:new).with(data)    { grid_properties }
    end

    specify do
      subject.update_grid_properties(data)

      expect(Google::Apis::SheetsV4::UpdateSheetPropertiesRequest).to have_received(:new)
        .with(fields: sheet_property_fields, properties: sheet_properties)
      expect(Google::Apis::SheetsV4::SheetProperties).to have_received(:new)
        .with(grid_properties: grid_properties, sheet_id: 1)
    end
  end

  describe '#update_cells' do
    let(:data) do
      [
        [
          { value: 'Date', background_color: '#b7b7b7', text_format: { bold: true }, horizontal_alignment: :center },
          { value: 'Developer', background_color: '#b7b7b7', text_format: { bold: true }, horizontal_alignment: :left }
        ],
        [
          { value: '09/14/2016', horizontal_alignment: :center, number_format: { type: :text } },
          { value: 'Some Author Name', horizontal_alignment: :left }
        ]
      ]
    end
    let(:updated_cells_reguest) { double :updated_cells_reguest }
    let(:grid_range)            { double :grid_range }
    let(:formatted_rows)        { double :formatted_rows }

    before do
      allow(Google::Apis::SheetsV4::UpdateCellsRequest).to receive(:new)
      allow(Google::Apis::SheetsV4::GridRange).to receive(:new) { grid_range }
      allow_any_instance_of(described_class).to receive(:format_rows_from).with(data) { formatted_rows }
    end

    specify do
      subject.update_cells(data)

      expect(Google::Apis::SheetsV4::UpdateCellsRequest).to have_received(:new)
        .with(rows: formatted_rows, fields: '*', range: grid_range) { updated_cells_reguest }
    end
  end

  describe '#merge_cells' do
    let(:data) do
      {
        type: :merge_all,
        range: {
          start_row_index:    1,
          end_row_index:      2,
          start_column_index: 0,
          end_column_index:   4
        }
      }
    end
    let(:grid_range)           { double :grid_range }
    let(:merged_cells_reguest) { double :merged_cells_reguest }

    before do
      allow_any_instance_of(described_class).to receive(:grid_range).with(data[:range]) { grid_range }
      allow(Google::Apis::SheetsV4::MergeCellsRequest).to receive(:new)
        .with(merge_type: data[:type], range: grid_range)
    end

    specify do
      subject.merge_cells(data)

      expect(Google::Apis::SheetsV4::MergeCellsRequest).to have_received(:new)
        .with(merge_type: data[:type], range: grid_range) { merged_cells_reguest }
    end
  end

  describe '#unmerge_cells' do
    let(:data) do
      {
        start_row_index:    0,
        end_row_index:      2,
        start_column_index: 0,
        end_column_index:   4
      }
    end
    let(:grid_range)             { double :grid_range }
    let(:unmerged_cells_request) { double :unmerged_cells_request }

    before do
      allow_any_instance_of(described_class).to receive(:grid_range).with(data) { grid_range }
      allow(Google::Apis::SheetsV4::UnmergeCellsRequest).to receive(:new) { unmerged_cells_request }
    end

    specify do
      subject.unmerge_cells(data)

      expect(Google::Apis::SheetsV4::UnmergeCellsRequest).to have_received(:new)
        .with(range: grid_range) { unmerged_cells_reguest }
    end
  end

  describe '#apply_changes!' do
    let!(:request_list)       { %w[request_1 request_2] }
    let(:batch_requests)      { double :batch_requests }
    let(:updated_spreadsheet) { double :updated_spreadsheet }

    before do
      subject.instance_variable_set(:@requests, request_list)
      allow(Google::Apis::SheetsV4::BatchUpdateSpreadsheetRequest).to receive(:new).with(requests: request_list)
                                                                                   .and_return(batch_requests)
      allow(service).to receive(:batch_update_spreadsheet).with('WE123', batch_requests).and_return(updated_spreadsheet)
    end

    specify do
      subject.apply_changes!

      expect(Google::Apis::SheetsV4::BatchUpdateSpreadsheetRequest).to have_received(:new).once
      expect(subject.instance_variable_get(:@requests)).to be_empty
    end
  end
end
