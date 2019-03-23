# frozen_string_literal: true

class CalendarEvent
  # Calendar fields from:
  # https://support.google.com/calendar/answer/37118?hl=en
  attr_accessor(:subject, :start_date, :end_date, :all_day_event, :description, :location, :private)
  attr_accessor(:class_category, :col_range, :row_range, :ref, :fill_color)

  def initialize(**options)
    %i[subject start_date end_date].each do |required_field|
      if options[required_field].nil?
        # Pry::ColorPrinter.pp(options)
        raise "Invalid value for #{required_field}: #{}"
      end
    end
    options.each do |key, value|
      instance_variable_set("@#{key}", value)
    end
  end

  def to_s
    "#{self.subject} - #{self.start_date} - #{self.end_date}"
  end

  def <=>(other)
    (self.start_date <=> other.start_date).zero? ? (self.end_date <=> other.end_date) : (self.start_date <=> other.start_date)
  end
end
