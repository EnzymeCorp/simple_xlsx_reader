frozen_string_literal: true

module SimpleXlsxReader
  class BufferLoader < Struct.new(:file) do
    attr_accessor :shared_strings, :sheet_parsers, :sheet_toc, :style_types, :base_date

    def init_sheets
      BufferZipReader.new(
        file: file,
        loader: self
      ).read

      sheet_toc.each_with_index.map do |(sheet_name, _sheet_number), i|
        # sheet_number is *not* the index into xml.sheet_parsers
        SimpleXlsxReader::Document::Sheet.new(
          name: sheet_name,
          sheet_parser: sheet_parsers[i]
        )
      end
    end

    BufferZipReader = ZipReader.new(:file, :loader, keyword_init: true) do
      attr_reader :zip

      def initialize(*args)
        super
        @zip = SimpleXlsxReader::Zip.open_buffer(buffer)
      end
    end
  end
end
