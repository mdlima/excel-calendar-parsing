require 'rubyXL'
# require 'rubyXL/convenience_methods/cell'
# require 'rubyXL/convenience_methods/color'
require 'loofah'
require_relative 'calendar_event'
require_relative 'calendar'

require 'pry-byebug'

# rubocop:disable Metrics/MethodLength, Metrics/AbcSize
def exec(source_file, dest_file)
  source_file ||= './Timetable Master Wintour UB 2019.xlsx'
  puts "Parsing #{source_file}"
  workbook = RubyXL::Parser.parse(source_file)
  worksheet = workbook.worksheets.first
  puts "Found #{workbook.worksheets.length} worksheets"

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
  # merged_ranges = worksheet.merged_cells.map { |mc| {mc.ref.row_range} }
  days = {}
  num_rows = 0

  worksheet.each do |row|
    all_day_event = false
    time = nil
    row = row.cells
    calendar_started ||= Date::MONTHNAMES.compact.map(&:capitalize).include?(row.first.value.to_s.capitalize) if row.first
    unless calendar_started
      puts 'Parsing class_categories'
      row.each do |cell|
        if cell && !class_categories[cell.fill_color] && cell.fill_color != 'ffffff'
          class_categories[cell.fill_color] = cell.value
          puts "#{cell.value} category for color #{cell.fill_color}"
        end
      end
      next
    end
    if row.compact.map(&:number_format).compact.map(&:format_code).include?('d-mmm')
      # puts 'Parsing Header row with dates'
      # Header row with dates
      days = {}
      row.each do |cell|
        next unless cell.value

        days[cell.r.first_col] = DateTime.new(DateTime.now.year, cell.value.month, cell.value.day, 0, 0, 0, DateTime.now.offset).to_time
      end
      next
    end

    row.compact.each do |cell|
      if cell.number_format && cell.number_format.format_code == 'h:mm'
        time = cell.raw_value.to_f
        next
      end

      next unless cell.value && cell.value.length > 1

      # puts "Parsing event #{cell.value} at cell #{cell.r}"
      event_text = Loofah.fragment(cell.value).scrub!(:prune).to_text(encode_special_chars: false)
      next if event_text.to_s.strip.empty?

      unless days[cell.r.first_col]
        puts "Cell text without corresponding date at cell #{cell.r}"
        # exit
        next
      end

      current_date = days[cell.r.first_col]
      # Find event end time
      start_time, end_time = nil
      event_text_match = event_text.match(/\A(.+)\s+(\d+(?:h|:)\d*)\s?-\s?(\d+(?:h|:)\d*)\s*(.*)/)
      if event_text_match && event_text_match.captures.length >= 3
        # Event has time information
        event_subject = event_text_match.captures.first
        start_time, end_time = event_text_match.captures[1..2].map { |t| Time.parse(t, current_date) }
        event_location = event_text_match.captures.length > 3 && !event_text_match.captures.last.to_s.strip.empty? ? event_text_match.captures.last : nil
      else
        # puts "Event text doesn't contain time information: event #{event_text} at cell #{cell.r}"
        # Try to estimate duration based on merged cell ranges
        cell_range = worksheet.merged_cells.filter { |mc| mc.ref.row_range.include?(cell.r.first_row) && mc.ref.col_range.include?(cell.r.first_col) }
        if cell_range.empty?
          puts "Couldn't estimate event time: event #{event_text} at cell #{cell.r}"
          next
        end

        event_subject = event_text
        if time
          # Event start time is known
          start_time = current_date + time
          end_time = start_time + cell_range.first.ref.row_range.size * 30.0 / 60.0 / 24.0
          event_location = nil
        else
          # Event start time is empty, must be a multi-day, all-day event
          all_day_event = true
          start_time = current_date
          end_time = days[cell.r.first_col + cell_range.last.ref.col_range.size - 1] + 1
        end
      end

      unless start_time
        puts "Event time not identified: event #{event_text} at cell #{cell.r}"
        next
      end

      calendar_events << CalendarEvent.new(
        subject: event_subject.strip,
        start_date: start_time,
        # start_time: start_time,
        end_date: end_time,
        # end_time: end_time,
        all_day_event: all_day_event,
        description: event_text.strip,
        location: event_location,
        private: false,
        class_category: class_categories[cell.fill_color],
        col_range: cell.r.col_range,
        row_range: cell.r.row_range,
        ref: cell.r.to_s,
        fill_color: cell.fill_color
      )
    end

    # row_cells = row.map(&:value)
    num_rows += 1
  end
  puts "Parsed total #{num_rows} rows"
  # binding.pry
  calendar_events.sort!
  dest_file ||= 'parsed.csv'
  calendar_events.to_csv(dest_file)
end

exec ARGV[0], ARGV[1]
# rubocop:enable Metrics/MethodLength, Metrics/AbcSize
