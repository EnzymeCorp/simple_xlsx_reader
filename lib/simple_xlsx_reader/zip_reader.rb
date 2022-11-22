frozen_string_literal: true

module SimpleXlsxReader
   class ZipReader < Struct.new(:file, :loader, keyword_init: true) do
    attr_reader :zip

    def initialize(*args)
      super
    end

    def read
      entry_at('xl/workbook.xml') do |file_io|
        loader.sheet_toc, loader.base_date = Loader::WorkbookParser.parse(file_io)
      end

      entry_at('xl/styles.xml') do |file_io|
        loader.style_types = Loader::StyleTypesParser.parse(file_io)
      end

      # optional feature used by excel,
      # but not often used by xlsx generation libraries
      if (ss_entry = entry_at('xl/sharedStrings.xml'))
        ss_entry.get_input_stream do |file|
          loader.shared_strings = Loader::SharedStringsParser.parse(file)
        end
      else
        loader.shared_strings = []
      end

      loader.sheet_parsers = []

      # Sometimes there's a zero-index sheet.xml, ex.
      # Google Docs creates:
      # xl/worksheets/sheet.xml
      # xl/worksheets/sheet1.xml
      # xl/worksheets/sheet2.xml
      # While Excel creates:
      # xl/worksheets/sheet1.xml
      # xl/worksheets/sheet2.xml
      add_sheet_parser_at_index(nil)

      i = 1
      i += 1 while add_sheet_parser_at_index(i)
    end

    def entry_at(path, &block)
      # Older and newer (post-mid-2021) RubyZip normalizes pathnames,
      # but unfortunately there is a time in between where it doesn't.
      # Rather than require a specific version, let's just be flexible.
      entry =
        zip.find_entry(path) || # *nix-generated
        zip.find_entry(path.tr('/', '\\')) || # Windows-generated
        zip.find_entry(path.downcase) || # Sometimes it's lowercase
        zip.find_entry(path.tr('/', '\\').downcase) # Sometimes it's lowercase

      if block
        entry.get_input_stream(&block)
      else
        entry
      end
    end

    def add_sheet_parser_at_index(index)
      sheet_file_name = "xl/worksheets/sheet#{index}.xml"
      return unless (entry = entry_at(sheet_file_name))

      parser =
        Loader::SheetParser.new(
          file_io: entry.get_input_stream,
          loader: loader
        )

      relationship_file_name = "xl/worksheets/_rels/sheet#{index}.xml.rels"
      if (rel = entry_at(relationship_file_name))
        parser.xrels_file = rel.get_input_stream
      end

      loader.sheet_parsers << parser
    end
  end

  def self.cast(value, type, style, options = {})
    return nil if value.blank?

    # Sometimes the type is dictated by the style alone
    if type.nil? ||
       (type == 'n' && %i[date time date_time].include?(style))
      type = style
    end

    casted =
      case type

      ##
      # There are few built-in types
      ##

      when 's' # shared string
        options[:shared_strings][value.to_i]
      when 'n' # number
        value.to_f
      when 'b'
        value.to_i == 1
      when 'str'
        value
      when 'inlineStr'
        value

      ##
      # Type can also be determined by a style,
      # detected earlier and cast here by its standardized symbol
      ##

      when :string, :unsupported
        value
      when :fixnum
        value.to_i
      when :float
        value.to_f
      when :percentage
        value.to_f / 100
      # the trickiest. note that  all these formats can vary on
      # whether they actually contain a date, time, or datetime.
      when :date, :time, :date_time
        value = Float(value)
        days_since_date_system_start = value.to_i
        fraction_of_24 = value - days_since_date_system_start

        # http://stackoverflow.com/questions/10559767/how-to-convert-ms-excel-date-from-float-to-date-format-in-ruby
        date = options.fetch(:base_date, DATE_SYSTEM_1900) + days_since_date_system_start

        if fraction_of_24.positive? # there is a time associated
          seconds = (fraction_of_24 * 86_400).round
          return Time.utc(date.year, date.month, date.day) + seconds
        else
          return date
        end
      when :bignum
        if defined?(BigDecimal)
          BigDecimal(value)
        else
          value.to_f
        end

      ##
      # Beats me
      ##

      else
        value
      end

    if options[:url]
      Hyperlink.new(options[:url], casted)
    else
      casted
    end
  end
end
