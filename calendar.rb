# frozen_string_literal: true

# require_relative 'calendar_event'

class Calendar < Array
  def <<(calendar_event)
    # Search for same event in a cell above, meaning that it's the same event spanning over multiple cells
    self.each do |ce|
      next unless ce.end_cell == [calendar_event.start_cell, [-1, 0]].transpose.map { |x| x.reduce(:+) }

      ce.end_cell = calendar_event.start_cell
      ce.end_time = calendar_event.end_time
      return self
    end

    super
  end

  def unique_id
    "#{self.subject} - #{self.start_date} - #{self.all_day_event ? '' : self.start_time}"
  end

  def to_s
    "#{self.subject} - #{self.start_date} - #{self.all_day_event ? '' : self.start_time}"
  end
end
