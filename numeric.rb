class Numeric
  def percent_of(n)
    (self.to_f / n.to_f * 100.0).round 2
  end

  def format_thousands()
    self.to_s.reverse.gsub(/...(?=.)/,'\&,').reverse
  end
end