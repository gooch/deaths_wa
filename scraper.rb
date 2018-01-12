# This is a template for a Ruby scraper on morph.io (https://morph.io)
# including some code snippets below that you should find helpful

require 'scraperwiki'
require 'csv'

current_data = './BITRE_ARDD_Fatal_Crashes_November_2017_Update_II.xlsx'
current_data = './Fatal_Crashes_December_2017.csv'
# current_data = './BITRE_ARDD_Fatalities_November_2017_Update_II.xlsx'

# 'Crash ID', 'State', 'Date', 'Month', 'Year', 'Dayweek', 'Time', 'Crash Type', 'Number of Fatalities', 'Bus Involvement', 'Rigid Truck Involvement', 'Articulated Truck Involvement', 'Speed Limit'

# crash data, xlsx
# 'Crash ID', 'State', 'Date', 'Month', 'Year', 'Dayweek', 'Time', 'Crash Type', '"Bus  Involvement"', '"Rigid Truck Involvement"', '"Articulated Truck  Involvement "', 'Speed Limit', 'Road User', 'Gender', 'Age'

# CSV headers
# CrashID,State,Date,Month,Year,Time,Crash_Type,Bus_Involvement,Heavy_Rigid_Truck_Involvement,Articulated_Truck_Involvement,Speed_Limit,Road_User,Gender,Age


# location of csv on bitre.gov.au
# #text > div:nth-child(6) > ul > li > ul > li:nth-child(4) > a

CSV.foreach(current_data, headers: true) do |row|
  record = {
    'Crash_id'                      => row[0],
    'State'                         => row[1],
    'Date'                          => row[2],
    'Time'                          => row[6],
    'Crash_Type'                    => row[7],
    'Fatalities'                    => row[8],
    'Bus_Involvement'               => row[9],
    'Rigid_Truck_Involvement'       => row[10],
    'Articulated_Truck_Involvement' => row[11],
    'Speed_Limit'                   => row[12],
#    'Road_User'                     => row[12],
#    'Gender'                        => row[13],
#    'Age'                           => row[14]
  }

  test_query = "* from data where `Crash_id` = '%d'" % [record['Crash_id']]

  if (ScraperWiki.select(test_query).empty? rescue true)
    puts "Saving new record #{record['Crash_id']} — #{record['Crash_Type']}, #{record['State']}"
    # ScraperWiki.save_sqlite(['Crash_id', 'Road_User', 'Gender', 'Age'], record)
    ScraperWiki.save_sqlite(['Crash_id'], record)
  else
    puts "Skipping already saved record " + record['Crash_id']
  end
end

#
#  test_query = "* from data where
#    `Crash_id` = '%d' and
#    `Road_User` = '%s' and 
#    `Gender` = '%s' and
#    `Age` = %d" % [record['Crash_id'], record['Road_User'], record['Gender'], record['Age']]
