# Google Docs

`google_docs` is a wrapper around [google-api-ruby-client](https://github.com/google/google-api-ruby-client).
It allows you to get information about the spreadsheet and format it.
Currently, gem consists of two classes `GoogleDocs::SpreadSheet` and `GoogleDocs::Sheet` and uses `Sheets API V4`.

Main features:
  - Set up columns width
  - Update grid properties of the sheet (freeze columns/rows)
  - Update and format cells
  - Merge/unmerge cells in the given range
  - Export google sheet documents as PDF
  - Read information from certain sheet

### Pre requirement
In order to use our gem your application should have configured [Google OmniAuth](https://github.com/google/google-api-ruby-client#authorization) already. Also you should configure [Google Sheet API](https://developers.google.com/sheets/api/quickstart/ruby)

### Installation
 You can add it to your Gemfile with:
 ```sh
gem 'google_docs'
```
Then run `bundle install`

### Getting started
Initialize new object of `GoogleDocs::SpreadSheet` class
```sh
spreadsheet = GoogleDocs::SpreadSheet.new('some_google_access_token', 'some_spreadsheet_id')
```
After successfull intialization we can call the following methods:
* `spreadsheet.properties` returns properties of the spreadsheet
* `spreadsheet.sheets` get all sheets of the spreadsheet.

If you want to get sheet directly by key:
 ```sh
sheet = GoogleDocs::Sheet.new(spreadsheet_id: required, sheet_id: required, google_access_token: required)
```
By default, after object initialization there are no modification request to be executed.
You should prepare requests and only then call `sheet.apply_changes!` to send requests.
To add requests you can run following methods:

* `sheet.setup_columns_width(data)` - build request to set up columns width. Argument `data` should have the following format:
    ```
    {
      'A'  => integer,
      'D'  => integer,
      ...
    }
    ```
    The keys of hash are the columns letters and values of hash - width in pixels for these names in sheet respectively.
    For instance, suppose, you have the following sheet: ![sheet_without_width](/samples/google_sheet_without_width.png)
    and our data are the following:
    ```
    {
      'A' => 50,
      'B' => 150,
      'C' => 150,
      'D' => 100,
      'E' => 50,
      'F' => 100,
      'G' => 700
    }
    ```
    After our method have been executed we receive:
    ![sheet_with_width](/samples/google_sheet_set_width.png)


*  `sheet.update_grid_properties(props)` - build request to update properties of the sheet, i.e. to freeze rows and columns while user is scrolling document. Argument `props` must have the following format:

    ```
    {
        row_count:           integer,
        column_count:        integer,
        frozen_row_count:    integer,
        frozen_column_count: integer,
        hide_gridlines:      boolean
    }
    ```

    For instance, suppose, you have the following data:
    ```
    props = {
      row_count:           7,
      column_count:        7,
      frozen_row_count:    2,
      frozen_column_count: 1,
      hide_gridlines:      true
    }
    ```
    After our method have been executed we receive:
    ![update_grid_properties](/samples/google_sheet_update_grid_properties.png)



* `sheet.update_cells(data)` - build request to update all cells in a range with new data. 
Argument `data` is an array with rows. Each row is an array of cells. Cell can be a String, Number, Boolean or Hash.
When the value of cell starts from '=' it will be sent as formula.
When cell is a Hash it can contain additional formatting options.
Schema of the `data` attribute:
    ```
    [
        [
            {
                value: 'any cell value'
                borders: {
                  left, right, top, bottom: {
                    style: :dotted, :dashed, :solid, :double
                    width: 1..3
                    color: '#RRGGBB'
                  }
                }
                background_color:       '#RRGGBB'
                horizontal_alignment:   :left, :center, :right
                vertical_alignment:     :top, :middle, :bottom
                hyperlink_display_type: :linked, :plain_text
                wrap_strategy:          :overflow_cell, :legacy_wrap, :clip, :wrap
                text_direction:         :left_to_right, :right_to_left
                number_format: {
                  type:    :text, :number, :percent, :currency, :date, :time, :datetime, :scientific
                  pattern: "#,##0.0000"
                }
                text_format: {
                  foreground_color: '#RRGGBB'
                  font_family:      'Terminus'
                  font_size:        integer
                  bold:             boolean
                  italic:           boolean
                  strikethrough:    boolean
                  underline:        boolean
                }
            },
            'string cell',
            number_cell
        ],
        # new row
        [ {cell 1}, {cell 2}, ... {cell n} ]
    ]
    ```

    For instance, data could look like:

    ```
    data = [
      [{ value: 'books', background_color: '#f1c232', text_format: { bold: true }, horizontal_alignment: :left,
         borders: { bottom: { style: :solid, width: 3, color: '#RRGGBB' },
                    right: { style: :solid, width: 3, color: '#RRGGBB' } } }
      ],
      [
        { value: 'id',           background_color: '#b7b7b7', text_format: { bold: true }, horizontal_alignment: :right },
        { value: 'author',       background_color: '#b7b7b7', text_format: { bold: true }, horizontal_alignment: :center },
        { value: 'title',        background_color: '#b7b7b7', text_format: { bold: true }, horizontal_alignment: :center },
        { value: 'genre',        background_color: '#b7b7b7', text_format: { bold: true }, horizontal_alignment: :center },
        { value: 'price',        background_color: '#b7b7b7', text_format: { bold: true }, horizontal_alignment: :center },
        { value: 'publish_date', background_color: '#b7b7b7', text_format: { bold: true }, horizontal_alignment: :center },
        { value: 'description',  background_color: '#b7b7b7', text_format: { bold: true }, horizontal_alignment: :center }
      ],
      [
        { value: 'bk101', text_format: { foreground_color: '#0143DF' }, horizontal_alignment: :right },
        { value: 'Gambardella, Matthew', text_format: { italic: true }, horizontal_alignment: :center },
        { value: "XML Developer's Guide", text_format: { underline: true }, number_format: { type: :text }, horizontal_alignment: :center },
        { value: 'Computer', text_format: { bold: true }, horizontal_alignment: :center },
        { value: 44.95, number_format: { type: :currency }, text_format: { strikethrough: true, foreground_color: '#FF0000' }, horizontal_alignment: :center },
        { value: '10/01/2000', number_format: { type: :date}, horizontal_alignment: :center },
        { value: 'An in-depth look at creating applications', text_format: { font_family: 'Courier New' }, horizontal_alignment: :left },
      ],
      [
        { value: 'bk102', text_format: { foreground_color: '#0143DF' }, horizontal_alignment: :right },
        { value: 'Ralls, Kim', text_format: { italic: true }, horizontal_alignment: :center },
        { value: "Midnight Rain", text_format: { underline: true }, number_format: { type: :text }, horizontal_alignment: :center },
        { value: 'Fantasy', text_format: { bold: true }, horizontal_alignment: :center },
        { value: 10.99, number_format: { type: :currency }, horizontal_alignment: :center },
        { value: '16/12/2000', number_format: { type: :date}, horizontal_alignment: :center },
        { value: 'A former architect battles corporate zombies, an evil sorceress', text_format: { font_family: 'Courier New' }, horizontal_alignment: :left },
      ]
    ]
    ```
    After our method have been executed we receive:
    ![update_cells](/samples/google_sheet_update_cells.png)


* `sheet.merge_cells(props)` - build request to merge cells in the given range. Index count starts from 0. Argument `props` has the following format:

    ```
    {
        type: :merge_all, :merge_columns, :merge_rows
        range: {
          start_row_index:    integer,
          end_row_index:      integer,
          start_column_index: integer,
          end_column_index:   integer
        }
    }
    ```
    For instance,
    ```
    props = {
      type: :merge_rows,
      range: {
        start_row_index:    0,
        end_row_index:      1,
        start_column_index: 0,
        end_column_index:   7
      }
    }
   ```
    After our method have been executed we receive:
    ![merge_cells](/samples/google_sheet_merge_cells.png)

* `sheet.unmerge_cells(range)` - build request to unmerge cells in the given range, where argument `range` has the following format:

    ```
    {
        start_row_index:    integer,
        end_row_index:      integer,
        start_column_index: integer,
        end_column_index:   integer
    }
    ```
* `sheet.apply_changes!` - execute all reguests.


* `sheet.get_rows_values`  - read information from each row in the sheet. Output example:
    ```
    [
      ["books"],
      ["id", "author", "title", "genre", "price", "publish_date", "description"],
      ["bk101", "Gambardella, Matthew", "XML Developer's Guide", "Computer", "44.95", "2000-10-01", "An in-depth look at creating applications"],
      ["bk102", "Ralls, Kim", "Midnight Rain", "Fantasy", "5.95", "2000-12-16", "A former architect battles corporate zombies, an evil sorceress, and her own childhood to become queen of the world"],
      ["bk103", "Corets, Eva", "Maeve Ascendant", "Fantasy", "5.95", "2000-11-17", "After the collapse of a nanotechnology society in England, the young survivors lay the foundation for a new society"],
      ["bk104", "Thurman, Paula", "Splish Splash", "Romance", "4.95", "2000-11-02", "A deep sea diver finds true love twenty thousand leagues beneath the sea"],
      ["bk105", "Kress, Peter", "Paradox Lost", "Science Fiction", "6.95", "2000-11-02", "After an inadvertant trip through a Heisenberg Uncertainty Device, James Salway discovers the problems of being quantum"]
    ]
    ```

* `sheet.download_pdf { |file| model.update(attachment: file) }` - export google sheet document as PDF 

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/jetruby/google_docs-ruby. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [Contributor Covenant](http://contributor-covenant.org) code of conduct.

## License

The gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).

## Code of Conduct

Everyone interacting in the ApolloUploadServer project’s codebases, issue trackers, chat rooms and mailing lists is expected to follow the [code of conduct](https://github.com/jetruby/google_docs-ruby/blob/master/CODE_OF_CONDUCT.md).

## About JetRuby
ApolloUploadServer is maintained and founded by JetRuby Agency.

We love open source software!
See [our projects][portfolio] or
[contact us][contact] to design, develop, and grow your product.

[portfolio]: http://jetruby.com/portfolio/
[contact]: http://jetruby.com/#contactUs
