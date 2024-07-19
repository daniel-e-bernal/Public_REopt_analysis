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
"""
Below are two function that can be used to produce random lat/long coordinates, however, some of them
may be generated and are in the ocean.
"""
# Create Random Generator function to produce random lat/long locations within the US
function generate_random_latitude()
    # Latitude bounds for contiguous US (approximate)
    min_lat, max_lat = 26.396308, 35.384358
    #creates a random number from 0 - 1 using rand() then scales it based on the upper and lower bounds
    #add the minimum bound of lat to ensure within US region
    lat = rand() * (max_lat - min_lat) + min_lat
   
    return lat
end

# Create Random Generator function to produce random lat/long locations within the US
function generate_random_longitude()
    
    # Longitude bounds for contiguous US (approximate)
    min_lon, max_lon = -125.0, -66.93457
    #creates a random number from 0 - 1 using rand() then scales it based on the upper and lower bounds
    #add the minimum bound of long to ensure within US region
    
    lon = rand() * (max_lon - min_lon) + min_lon
    
    return lon
end


"""
======================================================================================================================================================
"""
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
    elec_demand_charge = rand() * (max_dem - min_dem) + min_dem
    return elec_demand_charge
end

#random generator to create random cost per kW for PV 
function generate_random_PV_cost()
    #low bound and upper bound per kWh
    min_cost = 1700
    max_cost = 4200
    PV_cost = rand() * (max_cost - min_cost) + min_cost
    return PV_cost
end

#random generator to create random cost per kW for Wind
function generate_random_Wind_cost()
    #low bound and upper bound per kWh
    min_cost = 2386
    max_cost = 8000
    Wind_cost = rand() * (max_cost - min_cost) + min_cost
    return Wind_cost
end

#random generator to create random space (ft squared) constraints
function generate_random_roofspace()
    #low bound and upper bound per kWh
    min_space = 1000
    max_space = 1000000
    roof_space = rand() * (max_space - min_space) + min_space
    return roof_space
end

#random generator to create random land space in acres
function generate_random_landspace()
    #low bound and upper bound per kWh
    min_space = 0
    max_space = 10
    land_space = rand() * (max_space - min_space) + min_space
    return land_space
end

doe_reference_building_list = [ "FastFoodRest", "FullServiceRest", "Hospital", "LargeHotel",  "LargeOffice", "MediumOffice", "MidriseApartment", "Outpatient", "PrimarySchool",
"RetailStore", "SecondarySchool", "SmallHotel", "SmallOffice", "StripMall", "Supermarket", "Warehouse", "FlatLoad", "FlatLoad_24_5", "FlatLoad_16_7",
"FlatLoad_16_5", "FlatLoad_8_7", "FlatLoad_8_5" ]

#NEM list 
NEM_list = [0, 25, 50, 100, 200, 250, 300, 400, 500, 750, 1000, 1500, 2000, 2500, 3000, 3500, 4000, 4500, 5000, 10000]

#PV location
PV_location = ["both", "ground", "roof" ]

#read CSV file with USA coordinates
coord_data = CSV.read("./random lat and long/usa_coordinates 1.csv", DataFrame)
latitudes = coord_data.Latitude
longitudes = coord_data.Longitude

# Set up general inputs for randomized locations
data_file = "General_Inputs.json"
input_data = JSON.parsefile("scenarios/$data_file")
println("Successfuly parsed input data JSON file")

# up to 1000 runs
REopt_runs = fill(1, 2)

input_data_dic = [] #to store the input_data_site
site_analysis = [] #this is to store inputs and outputs of REopt runs
sites_iter = eachindex(REopt_runs)
for i in sites_iter
    input_data_site = copy(input_data)
    #Get lat/long for current site
    lat = latitudes[i]
    println("====================================")
    println("====================================")
    println(lat)
    lon = longitudes[i]
    println(lon)
    println("====================================")
    println("====================================")
    #Assign lat and lon to REopt run
    input_data_site["Site"]["latitude"] = lat
    input_data_site["Site"]["longitude"] = lon
    
    input_data_site["ElectricLoad"]["annual_kwh"] = generate_random_electricity_consumption()
    input_data_site["ElectricLoad"]["doe_reference_name"] = rand(doe_reference_building_list)
    input_data_site["ElectricTariff"]["blended_annual_demand_rate"] = generate_random_demand_charge()
    input_data_site["ElectricTariff"]["blended_annual_energy_rate"] = generate_random_electric_tariff()
    input_data_site["ElectricUtility"]["net_metering_limit_kw"] = rand(NEM_list)
    input_data_site["ElectricTariff"]["wholesale_rate"] = rand() * input_data_site["ElectricTariff"]["blended_annual_energy_rate"]
    input_data_site["PV"]["location"] = rand(PV_location)

    #if loop for space availability
    if input_data_site["PV"]["location"] == "ground"
        input_data_site["Site"]["land_acres"] = generate_random_landspace()
    elseif input_data_site["PV"]["location"] == "roof"
        input_data_site["Site"]["roof_squarefeet"] = generate_random_roofspace()
    else
        input_data_site["Site"]["roof_squarefeet"] = generate_random_roofspace()
        input_data_site["Site"]["land_acres"] = generate_random_landspace()
    end

    input_data_site["Wind"]["installed_cost_per_kw"] = generate_random_Wind_cost()
    input_data_site["PV"]["installed_cost_per_kw"] = generate_random_PV_cost()

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
    println("=======================================================")
    println("=======================================================")
    println("The input data site is")
    println(input_data_site)
    sleep(25)
    append!(site_analysis, [(input_data_site, results)])
    println("=======================================================")
    println("=======================================================")
    println("Completed Optimization Run #$i")
    println("=======================================================")
    println("=======================================================")
    #store input_data_site
    append!(input_data_dic, [deepcopy(input_data_site)])
    println(input_data_dic[i])
    sleep(10)
end
println("Completed Optimization")

#write onto JSON file
write.("./results/v2_REopt_data.json", JSON.json(site_analysis))
println("Successfully printed results on JSON file")

# Populate the DataFrame with the results produced and inputs
df = DataFrame(
    input_Latitude = [safe_get(input_data_dic[i], ["Site", "latitude"]) for i in sites_iter],
    input_Longitude = [safe_get(input_data_dic[i], ["Site", "longitude"]) for i in sites_iter],
    input_PV_location = [safe_get(input_data_dic[i], ["PV", "location"]) for i in sites_iter],
    input_PV_installed_cost = [round(safe_get(input_data_dic[i], ["PV", "installed_cost_per_kw"]), digits=2) for i in sites_iter],
    input_Wind_installed_cost = [round(safe_get(input_data_dic[i], ["Wind", "installed_cost_per_kw"]), digits=2) for i in sites_iter],
    input_Site_electric_load = [round(safe_get(input_data_dic[i], ["ElectricLoad", "annual_kwh"]), digits=0) for i in sites_iter],
    input_Site_building_type = [safe_get(input_data_dic[i], ["ElectricLoad", "doe_reference_name"]) for i in sites_iter],
    input_Site_roofspace = [round(safe_get(input_data_dic[i], ["Site", "roof_squarefeet"]), digits=0) for i in sites_iter],
    input_Site_landspace = [round(safe_get(input_data_dic[i], ["Site", "land_acres"]), digits=0) for i in sites_iter],
    input_Site_NEM_limit = [round(safe_get(input_data_dic[i], ["ElectricUtility", "net_metering_limit_kw"]), digits=0) for i in sites_iter],
    input_Site_net_billing_rate = [round(safe_get(input_data_dic[i], ["ElectricTariff", "wholesale_rate"]), digits=2) for i in sites_iter],
    input_Site_electricity_cost_per_kwh = [round(safe_get(input_data_dic[i], ["ElectricTariff", "blended_annual_energy_rate"]), digits=2) for i in sites_iter],
    input_Site_demand_charge_cost_per_kw = [round(safe_get(input_data_dic[i], ["ElectricTariff", "blended_annual_demand_rate"]), digits=2) for i in sites_iter],
    output_PV_size = [round(safe_get(site_analysis[i][2], ["PV", "size_kw"]), digits=0) for i in sites_iter],
    output_PV_energy_lcoe = [round(safe_get(site_analysis[i][2], ["PV", "lcoe_per_kwh"]), digits=0) for i in sites_iter],
    output_PV_energy_exported = [round(safe_get(site_analysis[i][2], ["PV", "annual_energy_exported_kwh"]), digits=0) for i in sites_iter],
    output_PV_energy_curtailed = [sum(safe_get(site_analysis[i][2], ["PV", "electric_curtailed_series_kw"], 0)) for i in sites_iter],
    output_Wind_size = [round(safe_get(site_analysis[i][2], ["Wind", "size_kw"]), digits=0) for i in sites_iter],
    output_Wind_energy_lcoe = [round(safe_get(site_analysis[i][2], ["Wind", "lcoe_per_kwh"]), digits=0) for i in sites_iter],
    output_Wind_energy_exported = [round(safe_get(site_analysis[i][2], ["Wind", "annual_energy_exported_kwh"]), digits=0) for i in sites_iter],
    output_Wind_energy_curtailed = [sum(safe_get(site_analysis[i][2], ["Wind", "electric_curtailed_series_kw"], 0)) for i in sites_iter],
    output_Grid_Electricity_Supplied_kWh_annual = [round(safe_get(site_analysis[i][2], ["ElectricUtility", "annual_energy_supplied_kwh"]), digits=0) for i in sites_iter],
    output_npv = [round(safe_get(site_analysis[i][2], ["Financial", "npv"]), digits=2) for i in sites_iter],
    output_lcc = [round(safe_get(site_analysis[i][2], ["Financial", "lcc"]), digits=2) for i in sites_iter]
    )
println(df)

# Define path to xlsx file
file_storage_location = "./results/REopt_data.xlsx"

# Check if the Excel file already exists
if isfile(file_storage_location)
    # Open the Excel file in read-write mode
    XLSX.openxlsx(file_storage_location, mode="rw") do xf
        counter = 0
        while true
            sheet_name = "v2_Results" * string(counter)
            try
                sheet = xf[sheet_name]
                counter += 1
            catch
                break
            end
        end
        sheet_name = "v2_Results" * string(counter)
        # Add new sheet
        XLSX.addsheet!(xf, sheet_name)
        # Write DataFrame to the new sheet
        XLSX.writetable!(xf[sheet_name], df)
    end
else
    # Write DataFrame to a new Excel file
    XLSX.writetable!(file_storage_location, df)
end

println("Successful write into XLSX file: $file_storage_location")