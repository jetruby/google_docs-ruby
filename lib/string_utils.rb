module StringUtils
  def self.camelize(string, uppercase_first_letter = true)
    string = if uppercase_first_letter
               string.sub(/^[a-z\d]*/) { $&.capitalize }
             else
               string.sub(/^(?:(?=\b|[A-Z_])|\w)/) { $&.downcase }
             end
    string.gsub(/(?:_|(\/))([a-z\d]*)/) { "#{Regexp.last_match(1)}#{Regexp.last_match(2).capitalize}" }.gsub('/', '::')
  end
end
