#!/usr/bin/env ruby
# Converts a spreadsheet of event information to an icalendar file.
# Spreadsheet columns should match arguments to createEvent().
# Usage: ruby csv2ical.rb spreadsheet.csv calendar.ics
# Original author: Jason McLaren <jason@fnord.ca>

require 'rubygems'
require 'csv'
require 'ri_cal'
require 'tzinfo'

# Arguments must be same as columns of spreadsheet
def createEvent(project_name, start_date, start_time, end_date, end_time,
  presenter, discipline, description, location, contact,
  hyperlink, francophone, free)
  event = RiCal.Event

  event.summary = project_name if project_name
  event.organizer = presenter if presenter
  event.location = location if location
  event.contact = contact if contact
  
  full_description = ""
  full_description << description if (description && !description.empty?)
  full_description << "\n" + hyperlink if (hyperlink && !hyperlink.empty?)
  full_description << "\nDiscipline: " + discipline if (discipline and !discipline.empty?)
  full_description << "\nEvent Francophone" if (francophone && francophone=='Y')
  full_description << "\nFree Event" if (free && free=='Y')
  event.description = full_description

  # ----- Guesstimate date and time -----
  # Skip events with no start date
  if (!start_date or start_date.empty?)
    puts 'Skipping event with no start date: ' + project_name
    return nil
  end
  # If no end date, make a one-day event
  if (!end_date or end_date.empty?)
    end_date = start_date
  end
  # If no end time, make an instantaneous event
  # TODO: make this a one-hour event if there is a start time
  if (!end_time or end_time.empty?)
    end_time = start_time
  end
  if (!start_time)
    start_time = ""
  end
  event.dtstart = DateTime.parse(start_date + " " + start_time, true).set_tzid("America/Vancouver")

  # If no start time, make an all-day event
  # TODO: try to get google to recognize this as an all-day event
  # by setting start/end Dates rather than DateTimes
  # or possibly with all-day extensions to iCal format
  if (start_time == "")
    end_dt = DateTime.parse(start_date, true).set_tzid("America/Vancouver") + 1
    event.dtend = end_dt
    event.dtend = event.dtend.set_tzid("America/Vancouver")
  else
    # Normal event!
    event.dtend = DateTime.parse(start_date + " " + end_time, true).set_tzid("America/Vancouver")
  end

  if (event.dtend.year != 2010 or event.dtstart.year != 2010)
    warning = "Warning: event not entirely in 2010: " + project_name
    warning << " start: " + event.dtstart.year.to_s + "/" + event.dtstart.month.to_s + "/" + event.dtstart.day.to_s
    warning << " end: " + event.dtend.year.to_s + "/" + event.dtend.month.to_s + "/" + event.dtend.day.to_s
    puts warning
  end

  # Handle recurring events
  if (end_date != start_date)
      end_dt = DateTime.parse(end_date, true).set_tzid("America/Vancouver")
      event.rrule = {:freq => "DAILY", :until => end_dt}
  end
    
  return event
end

# Read the csv file
if (ARGV.size != 2)
  puts "usage: " + $0 + " spreadsheet.csv calendar.ics"
  exit
end
filename = ARGV[0]
f = File.open(filename, 'r')
s = f.read
f.close
s.gsub!(/\r/, "\n")
rows = CSV.parse(s)
headers = rows.shift

# Create the calendar
row_num = 1
cal = RiCal.Calendar
rows.each do |row|
  row_num += 1
  if (row.length < 13)
    if (! row.select{|x| x}.empty?)
      puts "Ignoring incomplete row " + row_num.to_s
    end
  elsif (!row[0])
    puts "Ignoring incomplete row " + row_num.to_s
  else
    begin
      event = createEvent(*row)
    rescue
      event = nil
      puts "Bad date for event: " + row[0]
    end
    cal.add_subcomponent(event) if event
  end
end

# output
cal_filename = ARGV[1]
File.open(cal_filename, 'w') do |aFile|
  aFile.puts cal.export
end
