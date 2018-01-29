module GoogleDocs
  class SpreadSheet
    def initialize(google_access_token, spreadsheet_id)
      @service = SERVICE::SheetsService.new
      @service.authorization = google_access_token
      @spreadsheet = @service.get_spreadsheet(spreadsheet_id)
    end

    def sheets
      @spreadsheet.sheets.map do |sheet|
        Sheet.new(spreadsheet_id: @spreadsheet.spreadsheet_id, sheet_id: sheet.properties.sheet_id, service: @service)
      end
    end

    def properties
      @spreadsheet.properties
    end
  end
end
