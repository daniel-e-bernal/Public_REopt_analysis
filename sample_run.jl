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

# Function to safely extract values from JSON with default value if key is missing
# default must be an integer since it does not take in strings
function safe_get(data::Dict{String, Any}, keys::Vector{String}, default::Any=0)
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

#PV location
PV_location = ["both", "ground", "roof" ]

#lat long locations 
city = ["LasVegas", "Chicago"]
lat = [36.10, 41.8781]
long = [-115.1391, -87.6298]

# Set up general inputs for randomized locations
data_file = "General_Inputs.json"
input_data = JSON.parsefile("scenarios/$data_file")
println("Successfuly parsed input data JSON file")

# up to 1000 runs
REopt_runs = collect(1:4)

input_REopt_dic = [] #to store the inputs = REoptinputs(s)
input_data_dic = [] #to store the input_data_site
site_analysis = [] #this is to store inputs and outputs of REopt runs
println("Ready to start runs")
sites_iter = eachindex(REopt_runs)
for i in sites_iter
    input_data_site = copy(input_data)
    if sites_iter[i] < 3
        input_data_site["Site"]["latitude"] = lat[1]
        input_data_site["Site"]["longitude"] = long[1]
        input_data_site["ElectricLoad"]["city"] = city[1]
    else
        input_data_site["Site"]["latitude"] = lat[2]
        input_data_site["Site"]["longitude"] = long[2]
        input_data_site["ElectricLoad"]["city"] = city[2]
    end
    #making a scenario where Las Vegas only has roof space or only has land acres, and same with Chicago
    if sites_iter[i] == 1 || sites_iter[i] == 3
        input_data_site["PV"]["location"] = "ground"
        input_data_site["Site"]["land_acres"] = 1
        input_data_site["Site"]["roof_squarefeet"] = 0
        input_data_site["PV"]["array_type"] = 0 #0 is Ground Mount Fixed (Open Rack)
        #input_data_site["PV"]["tilt"] = 20
    else
        input_data_site["PV"]["location"] = "roof"
        input_data_site["Site"]["roof_squarefeet"] = 43560
        input_data_site["Site"]["land_acres"] = 0
        input_data_site["PV"]["array_type"] = 1 #1 is Rooftop, Fixed
        #input_data_site["PV"]["tilt"] = 20
    end

    #Electric Load and Tariff
    input_data_site["ElectricLoad"]["doe_reference_name"] = "FlatLoad"
    input_data_site["ElectricLoad"]["annual_kwh"] = 1000000
    input_data_site["ElectricTariff"]["wholesale_rate"] = 0
    input_data_site["ElectricTariff"]["blended_annual_energy_rate"] = 1.0
    input_data_site["ElectricTariff"]["blended_annual_demand_rate"] = 20

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

    #store inputs = REoptInputs(s) into the dictionary
    append!(input_REopt_dic, [deepcopy(inputs)])
    println("Copied inputs = REoptInputs(s) into the dictionary 'input_REopt_dic'")

    results = run_reopt([m1,m2], inputs)
    append!(site_analysis, [(input_data_site, results)])
    println("=======================================================")
    println("=======================================================")
    println("Completed Optimization Run #$i")
    println("=======================================================")
    println("=======================================================")
    #store input_data_site into the dictionary
    append!(input_data_dic, [deepcopy(input_data_site)])
    println("Copied inputs into the dictionary 'input_data_dic'")
end
println("Completed Optimization")

#write onto JSON file
write.("./results/PV_land_v_roof_REopt_results.json", JSON.json(site_analysis))
println("Successfully printed results on JSON file")
#write inputs onto JSON file
write.("./results/inputs_PV_land_v_roof.json", JSON.json(input_data_dic))
println("Successfully printed inputs onto JSON dictionary file")
#write inputs onto JSON file
write.("./results/inputs_REopt.json", JSON.json(input_REopt_dic))
println("Successfully printed REopt inputs onto JSON dictionary file")

# Set up extractable json file with all inputs to put onto DataFrame
inputs_file = "inputs_REopt.json"
inputs_all = JSON.parsefile("results/$inputs_file")
println("Successfuly parsed $inputs_file from JSON file")
    
# Populate the DataFrame with the results produced and inputs
df = DataFrame(
    city = [safe_get(input_data_dic[i], ["ElectricLoad", "city"]) for i in sites_iter],
    input_Latitude = [safe_get(input_data_dic[i], ["Site", "latitude"]) for i in sites_iter],
    input_Longitude = [safe_get(input_data_dic[i], ["Site", "longitude"]) for i in sites_iter],
    input_PV_location = [safe_get(input_data_dic[i], ["PV", "location"]) for i in sites_iter],
    input_PV_installed_cost = [round(inputs_all[i]["s"]["pvs"][1]["installed_cost_per_kw"], digits=2) for i in sites_iter],
    input_PV_kw_per_sqft = [round(inputs_all[i]["s"]["pvs"][1]["kw_per_square_foot"], digits=2) for i in sites_iter],
    input_PV_gcr = [round(inputs_all[i]["s"]["pvs"][1]["gcr"], digits=2) for i in sites_iter],
    input_PV_azimuth = [round(inputs_all[i]["s"]["pvs"][1]["azimuth"], digits=2) for i in sites_iter],
    input_PV_radius = [round(inputs_all[i]["s"]["pvs"][1]["radius"], digits=2) for i in sites_iter],
    input_PV_tilt = [round(inputs_all[i]["s"]["pvs"][1]["tilt"], digits=2) for i in sites_iter],
    input_PV_array_type = [round(inputs_all[i]["s"]["pvs"][1]["array_type"], digits=2) for i in sites_iter],
    input_PV_module_type = [round(inputs_all[i]["s"]["pvs"][1]["module_type"], digits=2) for i in sites_iter],
    input_Site_electric_load = [round(safe_get(input_data_dic[i], ["ElectricLoad", "annual_kwh"]), digits=0) for i in sites_iter],
    input_Site_building_type = [safe_get(input_data_dic[i], ["ElectricLoad", "doe_reference_name"]) for i in sites_iter],
    input_Site_roofspace = [round(safe_get(input_data_dic[i], ["Site", "roof_squarefeet"]), digits=0) for i in sites_iter],
    input_Site_landspace = [round(safe_get(input_data_dic[i], ["Site", "land_acres"]), digits=0) for i in sites_iter],
    input_Site_NEM_limit = [round(safe_get(input_data_dic[i], ["ElectricUtility", "net_metering_limit_kw"]), digits=0) for i in sites_iter],
    input_Site_net_billing_rate = [round(safe_get(input_data_dic[i], ["ElectricTariff", "wholesale_rate"]), digits=2) for i in sites_iter],
    input_Site_electricity_cost_per_kwh = [round(safe_get(input_data_dic[i], ["ElectricTariff", "blended_annual_energy_rate"]), digits=2) for i in sites_iter],
    input_Site_demand_charge_cost_per_kw = [round(safe_get(input_data_dic[i], ["ElectricTariff", "blended_annual_demand_rate"]), digits=2) for i in sites_iter],
    output_PV_size = [round(safe_get(site_analysis[i][2], ["PV", "size_kw"]), digits=0) for i in sites_iter],
    output_PV_yr1_production = [round(safe_get(site_analysis[i][2], ["PV", "year_one_energy_produced_kwh"]), digits=0) for i in sites_iter],
    output_PV_avg_annual_prod = [round(safe_get(site_analysis[i][2], ["PV", "annual_energy_produced_kwh"]), digits=0) for i in sites_iter],
    output_PV_energy_lcoe = [round(safe_get(site_analysis[i][2], ["PV", "lcoe_per_kwh"]), digits=0) for i in sites_iter],
    output_PV_energy_exported = [round(safe_get(site_analysis[i][2], ["PV", "annual_energy_exported_kwh"]), digits=0) for i in sites_iter],
    output_PV_energy_curtailed = [sum(safe_get(site_analysis[i][2], ["PV", "electric_curtailed_series_kw"], 0)) for i in sites_iter],
    output_Grid_Electricity_Supplied_kWh_annual = [round(safe_get(site_analysis[i][2], ["ElectricUtility", "annual_energy_supplied_kwh"]), digits=0) for i in sites_iter],
    output_npv = [round(safe_get(site_analysis[i][2], ["Financial", "npv"]), digits=2) for i in sites_iter],
    output_lcc = [round(safe_get(site_analysis[i][2], ["Financial", "lcc"]), digits=2) for i in sites_iter]
    )
println(df)

# Define path to xlsx file
file_storage_location = "./results/PV_land_v_roof_REopt_results.xlsx"

# Check if the Excel file already exists
if isfile(file_storage_location)
    # Open the Excel file in read-write mode
    XLSX.openxlsx(file_storage_location, mode="rw") do xf
        counter = 20
        while true
            sheet_name = "Results_" * string(counter)
            try
                sheet = xf[sheet_name]
                counter += 1
            catch
                break
            end
        end
        sheet_name = "Results_" * string(counter)
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