# This is a template for a Ruby scraper on morph.io (https://morph.io)
# including some code snippets below that you should find helpful

require 'scraperwiki'
require 'roo'

# current_data = './BITRE_ARDD_Fatal_Crashes_November_2017_Update_II.xlsx'
current_data = './BITRE_ARDD_Fatalities_November_2017_Update_II.xlsx'

# 'Crash ID', 'State', 'Date', 'Month', 'Year', 'Dayweek', 'Time', 'Crash Type', 'Number of Fatalities', 'Bus Involvement', 'Rigid Truck Involvement', 'Articulated Truck Involvement', 'Speed Limit'

# 'Crash ID', 'State', 'Date', 'Month', 'Year', 'Dayweek', 'Time', 'Crash Type', '"Bus  Involvement"', '"Rigid Truck Involvement"', '"Articulated Truck  Involvement "', 'Speed Limit', 'Road User', 'Gender', 'Age'

xlsx = Roo::Spreadsheet.open(current_data)

xlsx.default_sheet = '2011-2017'

last_row = xlsx.sheet('2011-2017').last_row

6.upto(last_row) do |row_n|
  row = xlsx.sheet('2011-2017').row(row_n)
  record = {
    'Crash_id'                      => row[0],
    'State'                         => row[1],
    'Date'                          => row[2],
    'Time'                          => row[6],
    'Crash_Type'                    => row[7],
    'Bus_Involvement'               => row[8],
    'Rigid_Truck_Involvement'       => row[9],
    'Articulated_Truck_Involvement' => row[10],
    'Speed_Limit'                   => row[11],
    'Road_User'                     => row[12],
    'Gender'                        => row[13],
    'Age'                           => row[14]
  }

  test_query = "* from data where
    `Crash_id` = '%d' and
    `Road_User` = '%s' and 
    `Gender` = '%s' and
    `Age` = %d" % [record['Crash_id'], record['Road_User'], record['Gender'], record['Age']]


  if (ScraperWiki.select(test_query).empty? rescue true)
    puts "Saving new record #{record['Crash_id']} — #{record['Crash_Type']}, #{record['State']}"
    ScraperWiki.save_sqlite(['Crash_id', 'Road_User', 'Gender', 'Age'], record)
  else
    puts "Skipping already saved record " + record['Crash_id']
  end
end
