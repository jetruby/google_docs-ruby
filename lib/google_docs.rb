require_relative 'google_docs/version'
require          'google/apis/sheets_v4'
require_relative 'google_docs/sheet'
require_relative 'google_docs/spread_sheet'

module GoogleDocs
  SERVICE = Google::Apis::SheetsV4
end
