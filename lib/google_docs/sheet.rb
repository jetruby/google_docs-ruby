require 'tempfile'
require_relative '../string_utils'

module GoogleDocs
  class Sheet
    LETTER_LIST = ('A'..'ZZZ').to_a.freeze
    # Set class instance variables
    # @param spreadsheet_id [String] id of spreadsheet
    # @param sheet_id [String] id of sheet
    # @param google_access_token [String]
    # @param service [Google::Apis::SheetsV4::SheetsService]
    # param google_access_token [String] google access token
    # @note if you have `service` you don't need `google_access_token`
    def initialize(spreadsheet_id:, sheet_id:, google_access_token: nil, service: nil)
      if service.nil?
        @service = SERVICE::SheetsService.new
        @service.authorization = google_access_token
      else
        @service = service
      end

      @spreadsheet = @service.get_spreadsheet(spreadsheet_id)
      @sheet       = @spreadsheet.sheets.find { |sheet| sheet.properties.sheet_id == sheet_id.to_i }
      @requests    = []

      raise ArgumentError, 'Invalid sheet id' if @sheet.nil?
    end

    # return all information from each row in sheet
    def get_rows_values
      last_column_letter = LETTER_LIST[@sheet.properties.grid_properties.column_count - 1]
      last_row_number    = @sheet.properties.grid_properties.row_count
      @service.get_spreadsheet_values(
        @spreadsheet.spreadsheet_id, "#{@sheet.properties.title}!A1:#{last_column_letter}#{last_row_number}"
      ).values
    end

    # create pdf file of sheet's content
    # @yield [file] Gives file to the block
    def download_pdf
      file = Tempfile.new(['sheet', '.pdf'])
      file.binmode
      file.write(pdf_request)
      yield(file)
      file.close
    end

    def pdf_request
      @service.http(:get, "https://docs.google.com/spreadsheets/d/#{@spreadsheet.spreadsheet_id}/export?exportFormat=pdf&gid=#{@sheet.properties.sheet_id}")
    end

    # @param data [Hash] keys are the columns names and values indicate desirable columns width in sheet
    # {
    #   {
    #     'A' => 'integer value',
    #     'D' => 'integer value',
    #     ...
    #   }
    # }
    # @note 'integer value' is positive pixel size for Google::Apis::SheetsV4::DimensionProperties
    def setup_columns_width(data)
      data.each do |key, value|
        append_request(
          update_dimension_properties: SERVICE::UpdateDimensionPropertiesRequest.new(
            range: SERVICE::DimensionRange.new(
              sheet_id:    @sheet.properties.sheet_id,
              dimension:   'COLUMNS',
              start_index: LETTER_LIST.index(key),
              end_index:   LETTER_LIST.index(key) + 1
            ),
            properties: SERVICE::DimensionProperties.new(pixel_size: value),
            fields:     'pixelSize'
          )
        )
      end
      true
    end

    # @param props [Hash]
    # @example of props value
    #   https://developers.google.com/sheets/reference/rest/v4/spreadsheets#GridProperties
    #   {
    #     row_count:           integer
    #     column_count:        integer
    #     frozen_row_count:    integer
    #     frozen_column_count: integer
    #     hide_gridlines:      bool
    #   }
    def update_grid_properties(props)
      append_request(
        update_sheet_properties: SERVICE::UpdateSheetPropertiesRequest.new(
          fields:     props.keys.map { |field| "gridProperties.#{StringUtils.camelize(field.to_s)}" }.join(','),
          properties: SERVICE::SheetProperties.new(
            grid_properties: SERVICE::GridProperties.new(props),
            sheet_id:        @sheet.properties.sheet_id
          )
        )
      )
    end

    # @param data [Array<Array>]
    # @example of data value
    #   [
    #     [ 'text', 'text', ...],                # first row
    #     [],                                    # second row (empty)
    #     [ 'text', { value: 'text', ...props }] # third row
    #     ...
    #   ]
    def update_cells(data)
      append_request(
        update_cells: SERVICE::UpdateCellsRequest.new(
          rows:   format_rows_from(data),
          fields: '*',
          range:  grid_range
        )
      )
    end

    # @param props [Hash]
    # @example of props value
    #   {
    #     type: :merge_all, :merge_columns, :merge_rows
    #     range: {
    #       start_row_index:    integer
    #       end_row_index:      integer
    #       start_column_index: integer
    #       end_column_index:   integer
    #     }
    #   }
    def merge_cells(props)
      append_request(
        merge_cells: SERVICE::MergeCellsRequest.new(
          merge_type: props[:type],
          range:      grid_range(props[:range])
        )
      )
    end

    # @param range [Hash]
    # @example of props value
    #   {
    #     start_row_index:    integer
    #     end_row_index:      integer
    #     start_column_index: integer
    #     end_column_index:   integer
    #   }
    def unmerge_cells(range)
      append_request(
        unmerge_cells: SERVICE::UnmergeCellsRequest.new(
          range: grid_range(range)
        )
      )
    end

    def apply_changes!
      @service.batch_update_spreadsheet(
        @spreadsheet.spreadsheet_id, SERVICE::BatchUpdateSpreadsheetRequest.new(requests: @requests)
      )
      @requests = []
    end

    private

    def append_request(request_params)
      @requests << SERVICE::Request.new(request_params)
    end

    def grid_range(range = {})
      SERVICE::GridRange.new(range.merge(sheet_id: @sheet.properties.sheet_id))
    end

    def format_rows_from(values)
      values.map do |row_cells|
        row_data = row_cells.map do |cell|
          if cell.is_a? Hash
            SERVICE::CellData.new(
              user_entered_value: extended_value_for(cell[:value]),
              user_entered_format: format_cell_with(cell.except(:value))
            )
          else
            SERVICE::CellData.new(user_entered_value: extended_value_for(cell))
          end
        end
        SERVICE::RowData.new(values: row_data)
      end
    end

    def detect_type_for(value)
      case value
      when String
        value.start_with?('=') ? :formula_value : :string_value
      when Integer, Float
        :number_value
      when TrueClass, FalseClass
        :bool_value
      else
        :string_value
      end
    end

    def extended_value_for(value)
      SERVICE::ExtendedValue.new(detect_type_for(value || value.to_s) => (value || value.to_s))
    end

    #
    # {
    #   borders: {
    #     left, right, top, bottom: {
    #       style: :dotted, :dashed, :solid, :double
    #       width: 1..3
    #       color: '#RRGGBB'
    #     }
    #   }
    #   background_color:       '#RRGGBB'
    #   horizontal_alignment:   :left, :center, :right
    #   vertical_alignment:     :top, :middle, :bottom
    #   hyperlink_display_type: :linked, :plain_text
    #   wrap_strategy:          :overflow_cell, :legacy_wrap, :clip, :wrap
    #   text_direction:         :left_to_right, :right_to_left
    #   number_format: {
    #     type:    :text, :number, :percent, :currency, :date, :time, :datetime, :scientific
    #     pattern: "#,##0.0000"
    #   }
    #   text_format: {
    #     foreground_color: '#RRGGBB'
    #     font_family:      'Terminus'
    #     font_size:        integer
    #     bold:             bool
    #     italic:           bool
    #     strikethrough:    bool
    #     underline:        bool
    #   }
    # }

    def format_cell_with(params)
      params = params.dup
      params[:background_color] &&= format_color   params[:background_color]
      params[:borders]          &&= format_borders params[:borders]
      params[:text_format]      &&= format_text    params[:text_format]
      params[:number_format]    &&= format_number  params[:number_format]
      SERVICE::CellFormat.new(params)
    end

    def format_text(props)
      props = props.dup
      props[:foreground_color] &&= format_color(props[:foreground_color])
      SERVICE::TextFormat.new(props)
    end

    def format_number(props)
      SERVICE::NumberFormat.new(props)
    end

    def format_color(hex)
      red, green, blue = hex_to_rgb(hex)
      SERVICE::Color.new(red: red, blue: blue, green: green, alpha: 1.0)
    end

    def format_borders(borders)
      SERVICE::Borders.new(borders.map { |position, props| [position, format_border(props)] }.to_h)
    end

    def format_border(props)
      SERVICE::Border.new(
        style: props[:style],
        width: props[:width],
        color: format_color(props[:color])
      )
    end

    # '#RRGGBB' => [0..1, 0..1, 0..1]

    def hex_to_rgb(hex)
      hex.scan(/[^#]{2}/).map { |color| (color.to_i(16).to_f / 255) }
    end
  end
end
