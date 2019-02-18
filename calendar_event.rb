# frozen_string_literal: true

class CalendarEvent
  # Calendar fields from:
  # https://support.google.com/calendar/answer/37118?hl=en
  attr_accessor(:subject, :start_date, :end_date, :all_day_event, :description, :location, :private)
  attr_accessor(:class_category, :start_cell, :end_cell, :style)

  def initialize(**options)
    options.each do |key, value|
      instance_variable_set("@#{key}", value)
    end
  end

  def unique_id
    "#{self.subject} - #{self.start_date} - #{self.all_day_event ? '' : self.start_time}"
  end

  def to_s
    "#{self.subject} - #{self.start_date} - #{self.all_day_event ? '' : (self.start_time + ' to ' + self.end_time)}"
  end
end
