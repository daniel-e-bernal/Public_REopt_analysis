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
using Shapefile
using LibGEOS
using GeoInterface
using ArchGDAL
using GeometryBasics

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
    min_lat, max_lat = 24.396308, 49.384358
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
Below is a new way to generate random lat/long variables depending on the SHP file that I uploaded to this.
"""

# Load the shapefile containing the boundaries of the continental US
function load_us_shapefile(filepath)
    try
        shapefile = Shapefile.Table(filepath)
        return shapefile
    catch e
        println("Error loading shapefile: $e")
        return nothing
    end
end
println("Successfully loaded shpfile")

# usage to load the shapefile
shapefile_path = "./ShapeFiles for US/International Boundary/tl_2023_us_internationalboundary.shp"
us_shapefile = load_us_shapefile(shapefile_path)

if us_shapefile === nothing
    error("Shapefile could not be loaded.")
else
    println("Shapefile loaded successfully.")
end

# Generate random latitude within the continental US bounds
function generate_random_latitude()
    min_lat, max_lat = 24.396308, 49.384358
    return rand() * (max_lat - min_lat) + min_lat
end
println("Successfully loaded latitude")
sleep(5)
# Generate random longitude within the continental US bounds
function generate_random_longitude()
    min_lon, max_lon = -125.0, -66.93457
    return rand() * (max_lon - min_lon) + min_lon
end
println("Successfully loaded longitude")
sleep(5)
"""
# Check if the location (lat/long) is within the US boundary
function is_within_us_boundary(lat, lon, shapefile)
    point = LibGEOS.createPoint(lon, lat)
    for feature in shapefile
        geom = Shapefile.shape(feature)
        poly = GeoInterface.coordinates(geom)
        #convert to LibGEOS Polygon
        polygon = LibGEOS.createPolygon(poly)
        if LibGEOS.contains(polygon, point)
            return true
        end
    end
    return false
end

# Generate a random valid location within the continental US
function generate_random_location(shapefile)
    while true
        lat = generate_random_latitude()
        lon = generate_random_longitude()
        if is_within_us_boundary(lat, lon, shapefile)
            return lat, lon
        end
    end
end
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
REopt_runs = fill(1, 3)

site_analysis = [] #this is to store inputs and outputs of REopt runs
sites_iter = eachindex(REopt_runs)
for i in sites_iter
    input_data_site = copy(input_data)

    #Get lat/long for current site
    lat = latitudes[i]
    lon = longitudes[i]
    #Assign lat and lon to REopt run
    input_data_site["Site"]["latitude"] = lat
    input_data_site["Site"]["longitude"] = lon

    #Site Specific Randomnization
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
    append!(site_analysis, [(input_data_site, results)])
    println("Completed Optimization Run #$i")
end
println("Completed Optimization")

# Populate the DataFrame with the results produced and inputs
df = DataFrame(
    input_Latitude = [safe_get(site_analysis[i][2], ["Site", "latitude"]) for i in sites_iter],
    input_Longitude = [safe_get(site_analysis[i][2], ["Site", "longitude"]) for i in sites_iter],
    input_PV_location = [safe_get(site_analysis[i][2], ["PV", "location"]) for i in sites_iter],
    input_PV_installed_cost = [round(safe_get(site_analysis[i][2], ["PV", "installed_cost_per_kw"]), digits=2) for i in sites_iter],
    input_Wind_installed_cost = [round(safe_get(site_analysis[i][2], ["Wind", "installed_cost_per_kw"]), digits=2) for i in sites_iter],
    input_Site_electric_load = [round(safe_get(site_analysis[i][2], ["ElectricLoad", "annual_kwh"]), digits=0) for i in sites_iter],
    input_Site_building_type = [safe_get(site_analysis[i][2], ["ElectricLoad", "doe_reference_name"]) for i in sites_iter],
    input_Site_roofspace = [round(safe_get(site_analysis[i][2], ["Site", "roof_squarefeet"]), digits=0) for i in sites_iter],
    input_Site_landspace = [round(safe_get(site_analysis[i][2], ["Site", "land_acres"]), digits=0) for i in sites_iter],
    input_Site_NEM_limit = [round(safe_get(site_analysis[i][2], ["ElectricUtility", "net_metering_limit_kw"]), digits=0) for i in sites_iter],
    input_Site_net_billing_rate = [round(safe_get(site_analysis[i][2], ["ElectricTariff", "wholesale_rate"]), digits=2) for i in sites_iter],
    input_Site_electricity_cost_per_kwh = [round(safe_get(site_analysis[i][2], ["ElectricTariff", "blended_annual_energy_rate"]), digits=2) for i in sites_iter],
    input_Site_demand_charge_cost_per_kw = [round(safe_get(site_analysis[i][2], ["ElectricTariff", "blended_annual_demand_rate"]), digits=2) for i in sites_iter],
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
file_storage_location = "C:/Users/dbernal/Documents/GitHub/Public_REopt_analysis/results/REopt_data.xlsx"

# Check if the Excel file already exists
if isfile(file_storage_location)
    # Open the Excel file in read-write mode
    XLSX.openxlsx(file_storage_location, mode="rw") do xf
        counter = 0
        while true
            sheet_name = "Results" * string(counter)
            try
                sheet = xf[sheet_name]
                counter += 1
            catch
                break
            end
        end
        sheet_name = "Results" * string(counter)
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