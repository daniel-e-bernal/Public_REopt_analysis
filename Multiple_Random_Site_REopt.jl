"""
This code is meant to create an excel file that contains randomized REopt runs from random locations within US boundaries.
The results will be output onto an excel sheet alongside key inputs for those outputs.

    Outputs of interest:
    NPV
    Technology size (PV and/or Wind)
    Others to be considered
"""

using REopt
using HiGHS
using JSON
using JuMP
using CSV 
using DataFrames #to construct comparison
using XLSX 
using DelimitedFiles
using Plots
using Dates
using Random

# Function to safely extract values from JSON with default value if key is missing
# default must be an integer since it does not take in strings
function safe_get(data::Dict{String, Any}, keys::Vector{String}, default=0)
    try
        for k in keys
            data = data[k]
        end
        return data
    catch e
        if e isa KeyError
            return default
        else
            rethrow(e)
        end
    end
end

# Create Random Generator function to produce random lat/long locations within the US
function generate_random_latitude()
    # Latitude bounds for contiguous US (approximate)
    min_lat, max_lat = 24.396308, 49.384358
    #creates a random number from 0 - 1 using rand() then scales it based on the upper and lower bounds
    #add the minimum bound of lat to ensure within US region
    lat = rand() * (max_lat - min_lat) + min_lat
   
    return lat, lon
end

# Create Random Generator function to produce random lat/long locations within the US
function generate_random_longitude()
    
    # Longitude bounds for contiguous US (approximate)
    min_lon, max_lon = -125.0, -66.93457
    #creates a random number from 0 - 1 using rand() then scales it based on the upper and lower bounds
    #add the minimum bound of long to ensure within US region
    
    lon = rand() * (max_lon - min_lon) + min_lon
    
    return lat, lon
end

# Create Random Generator function to produce random annual energy consumptions for electricity
function generate_random_electricity_consumption()
    #low bound for yearly energy consumption
    min_elec_energy = 100000
    #upper bound for yearly energy consumption
    max_elec_energy = 5000000
    #creates a random number from 0 - 1 using rand() then scales it based on the upper and lower bounds
    #add the minimum amount of annual energy consumption to ensure above lower bound
    elec_consumption = rand() * (max_elec_energy - min_elec_energy) + min_elec_energy
    return elec_consumption
end

#random generator to create random elec tariff for blended annual average
function generate_random_electric_tariff()
    #low bound and upper bound per kWh
    min_tariff = 0.01
    max_tariff = 0.15
    elec_tariff = rand() * (max_tariff - min_tariff) + min_tariff
    return elec_tariff
end

#random generator to create random elec demand charge for blended annual average
function generate_random_demand_charge()
    #low bound and upper bound per kWh
    min_dem = 15.00
    max_dem = 250.00
    elec_tariff = rand() * (max_dem - min_dem) + min_dem
    return elec_demand_charge
end

doe_reference_building_list = [ "FastFoodRest", "FullServiceRest", "Hospital", "LargeHotel",  "LargeOffice", "MediumOffice", "MidriseApartment", "Outpatient", "PrimarySchool",
"RetailStore", "SecondarySchool", "SmallHotel", "SmallOffice", "StripMall", "Supermarket", "Warehouse", "FlatLoad", "FlatLoad_24_5", "FlatLoad_16_7",
"FlatLoad_16_5", "FlatLoad_8_7", "FlatLoad_8_5" ]

#NEM list 
NEM_list = [0, 25, 50, 100, 200, 250, 300, 400, 500, 750, 1000, 1500, 2000, 2500, 3000, 3500, 4000, 4500, 5000, 10000]

#PV location
PV_location = ["both", "ground", "roof" ]

# Generate and print a random location
random_location = generate_random_location()
println("This is a Random Location test run: Latitude: ", random_location[1], ", Longitude: ", random_location[2])

# Set up general inputs for randomized locations
data_file = "General_Inputs.json"
input_data = JSON.parsefile("scenarios/$data_file")

# up to 1000 runs
REopt_runs = fill(1, 1000)

site_analysis = [] #this is to store inputs and outputs of REopt runs
sites_iter = eachindex(REopt_runs)
for i in sites_iter
    input_data_site = copy(input_data)
    #Site Specific Randomnization
    input_data_site["Site"]["latitude"] = generate_random_latitude()
    input_data_site["Site"]["longitude"] = generate_random_longitude()
    input_data_site["ElectricLoad"]["annual_kwh"] = generate_random_electricity_consumption()
    input_data_site["ElectricLoad"]["doe_reference_name"] = rand(doe_reference_building_list)
    input_data_site["ElectricTariff"]["blended_annual_demand_rate"] = generate_random_demand_charge()
    input_data_site["ElectricTariff"]["blended_annual_energy_rate"] = generate_random_electric_tariff()
    input_data_site["ElectricUtility"]["net_metering_limit_kw"] = rand(NEM_list)
    input_data_site["PV"]["location"] = rand(PV_location)
    
    s = Scenario(input_data_site)
    inputs = REoptInputs(s)

     # HiGHS solver
     m1 = Model(optimizer_with_attributes(HiGHS.Optimizer, 
     "time_limit" => 600.0,
     "mip_rel_gap" => 0.01,
     "output_flag" => false, 
     "log_to_console" => false)
     )

    m2 = Model(optimizer_with_attributes(HiGHS.Optimizer, 
     "time_limit" => 600.0,
     "mip_rel_gap" => 0.01,
     "output_flag" => false, 
     "log_to_console" => false)
     )            

    results = run_reopt([m1,m2], inputs)
    append!(site_analysis, [(input_data_site, results)])
end
println("Completed Optimization")