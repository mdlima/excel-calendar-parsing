require 'roo'
require 'loofah'
require_relative 'calendar_event'
require_relative 'calendar'

require 'pry-byebug'

def get_row(sheet, max_rows)
  last_row = nil
  sheet.each_row_streaming(max_rows: max_rows) do |row|
    last_row = row
  end
  last_row
end

# rubocop:disable Metrics/MethodLength, Metrics/AbcSize
def exec
  workbook = Roo::Spreadsheet.open('./Timetable Master Wintour UB 2019.xlsx', expand_merged_ranges: true)
  worksheets = workbook.sheets
  puts "Found #{worksheets.count} worksheets"

  # worksheets.each do |worksheet|
  #   puts "Reading: #{worksheet}"
  #   num_rows = 0
  #   workbook.sheet(worksheet).each_row_streaming do |row|
  #     row_cells = row.map { |cell| cell.value }
  #     num_rows += 1
  #   end
  #   puts "Read #{num_rows} rows"
  # end

  calendar_events = Calendar.new
  class_categories = {}
  calendar_started = false
  time = nil
  days = {}
  num_rows = 0

  workbook.sheet(worksheets.first).each_row_streaming do |row|
    calendar_started ||= Date::MONTHNAMES.compact.map(&:capitalize).include?(row.first.value.to_s.capitalize) if row.first
    unless calendar_started
      puts 'Parsing class_categories'
      row.each do |cell|
        class_categories[cell.style] = cell.value if cell.value
      end
      next
    end
    if row.map(&:type).include?(:date)
      puts 'Parsing Header row with dates'
      # Header row with dates
      days = {}
      row.each do |cell|
        next unless cell.value

        days[cell.coordinate.last] = {
          day: cell.value.day,
          month: cell.value.month,
        }
      end
      next
    end

    if row.first && row.first.type == :time
      time = row.first.to_s
      row[1..-1].each do |cell|
        next unless cell.value

        # puts "Parsing event #{cell.value} at cell #{cell.coordinate}"
        event_text = Loofah.fragment(cell.value).scrub!(:prune).text
        next if event_text.to_s.strip.empty?

        unless days[cell.coordinate.last]
          puts "Cell text without corresponding date at cell #{cell.coordinate}"
          # exit
          next
        end

        # Find event end time
        event_text_match = event_text.match(/\A(.+)\s+(\d+(?:h|:)\d*)\s?-\s?(\d+(?:h|:)\d*)\s*(.*)/)
        unless event_text_match && event_text_match.captures.length >= 3
          puts "Event text doesn't contain time information: event #{event_text} at cell #{cell.coordinate}"
          next
        end
        current_date = DateTime.new(DateTime.now.year, days[cell.coordinate.last][:month], days[cell.coordinate.last][:day])
        start_time, end_time = event_text_match.captures[1..2].map { |t| Time.parse(t, current_date) }

        calendar_events << CalendarEvent.new(
          subject: event_text_match.captures.first,
          start_date: start_time,
          # start_time: start_time,
          end_date: end_time,
          # end_time: end_time,
          all_day_event: false,
          description: event_text,
          # location: ,
          private: false,
          class_category: class_categories[cell.style],
          start_cell: cell.coordinate,
          end_cell: cell.coordinate,
          style: cell.style
        )
        calendar_events.last.location = event_text_match.captures.last if event_text_match.captures.length > 3 && !event_text_match.captures.last.to_s.strip.empty?
      end
    end
    # row_cells = row.map(&:value)
    num_rows += 1
  end
  puts "Parsed total #{num_rows} rows"
  binding.pry
end

exec
# rubocop:enable Metrics/MethodLength, Metrics/AbcSize
