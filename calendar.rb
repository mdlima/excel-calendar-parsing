# frozen_string_literal: true

require 'csv'
require_relative 'calendar_event'

class Calendar < Array
  # def <<(calendar_event)
  #   # Search for same event in a cell above, meaning that it's the same event spanning over multiple cells
  #   self.each do |ce|
  #     next unless ce.end_cell == [calendar_event.start_cell, [-1, 0]].transpose.map { |x| x.reduce(:+) }

  #     ce.end_cell = calendar_event.start_cell
  #     ce.end_time = calendar_event.end_time
  #     return self
  #   end

  #   super
  # end

  def to_csv(filename)
    CSV.open(filename, 'wb') do |csv|
      # header
      csv << ['Subject', 'Start Date', 'Start Time', 'End Date', 'End Time', 'All Day Event', 'Description', 'Location', 'Private']
      self.each do |event|
        csv << [
          event.subject,
          event.start_date.strftime('%F'),
          event.start_date.strftime('%R'),
          event.end_date.strftime('%F'),
          event.end_date.strftime('%R'),
          event.all_day_event,
          event.description,
          event.location,
          false
        ]
      end
    end
  end
end
