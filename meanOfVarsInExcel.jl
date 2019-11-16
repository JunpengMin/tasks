#This file shows how to read in excel files, do some computations and then export the results to one single excel files. Suppose there are 3 excel files describing the GDP, R&D expenditure, and labor panel data of three countries. Each file has 3 sheets and each sheet contains a panel data.To show it graphically:
#excel file 1: country1 contains GDP panel, R&D panel and labor panel, with rows as observation in each province and columns as data in each year and the size of each sheet is N × T
#excel file 2: country2 contains GDP panel, R&D panel and labor panel, with rows as observation in each province and columns as data in each year and the size of each sheet is N × T
#excel file 3: country3 contains GDP panel, R&D panel and labor panel, with rows as observation in each province and columns as data in each year and the size of each sheet is N × T
#
#For each file, we want to compute the mean of each variable(corresponding to each sheet) over provinces, thus for each sheet we obtain a time series data describing the variable level of that country. Then we want to juxtapose the three ts data to create a new sheet, which tells three ts data of that country. After doing this for those three countries, we then combining those three sheets into one single excel file. 

using XLSX
using DataFrames
using Statistics #used for computing row means
using CSV #for exporting to csv file, but unnecessary if to export to a single excel file

function genPanleSample(N, T)
    sample = rand(N, T)
    sampleDf = DataFrame(sample)
    names!(sampleDf, [Symbol("y$(1990+i)") for i in 1:T])
    insertcols!(sampleDf, 1, :id=>["$i" for i in 1:N])
    return sampleDf
end


"""
dfArray2Excel(dfArray,excelFileName)
Convert an array of dataframes to an excel, with each sheet containing one dataframe. Note that the excelFileName must be created first manually with at least as many sheets as dataframes in dfArray.
# Examples
```jldoctest
julia> meanOfDfArray(dfArray)
8
```
"""
function dfArray2Excel(dfArray, excelFileName)
	XLSX.openxlsx(excelFileName, mode="rw") do excelFile
		for i in 1:length(dfArray)
			sheeti = excelFile[i]
			XLSX.rename!(sheeti, "sheet$i")
			df = dfArray[i]
			columns = collect(eachcol(df))
			labels = String.(names(df))
			XLSX.writetable!(sheeti, columns, labels, anchor_cell=XLSX.CellRef("A1"))
		end	
	end
end

country1ArrayDf = [genPanleSample(31, 20), genPanleSample(31, 20), genPanleSample(31, 20)]
country2ArrayDf = [genPanleSample(31, 20), genPanleSample(31, 20), genPanleSample(31, 20)]
country3ArrayDf = [genPanleSample(31, 20), genPanleSample(31, 20), genPanleSample(31, 20)]

dfArray2Excel(country1ArrayDf, "country1.xlsx")
dfArray2Excel(country2ArrayDf, "country2.xlsx")
dfArray2Excel(country3ArrayDf, "country3.xlsx")


"""
csvSheets2DfArray(csvFilePath)
convert a excel file with mutiple sheets to a julia 1D Array of DataFrame

# Examples
```jldoctest
julia> csvSheets2ArrayOfDfs(csvFilePath)
8
```
"""
function excelSheets2DfArray(excelFilePath)
    file =  XLSX.readxlsx(excelFilePath)
    sheetNames = [i.name for i in file.workbook.sheets]
    dfArray = Array{DataFrame,1}(undef, length(sheetNames))    
    
    for (i, sheetname) in enumerate(sheetNames)
        sheet = file[sheetname]
        sheetMatrix = sheet[:]
        sheetVarNames = sheetMatrix[1, 2:end]
        sheetObsNames = sheetMatrix[2:end, 1]
        df = DataFrame(convert(Array{Float64,2}, sheetMatrix[2:end, 2:end]), Symbol.(sheetVarNames))
        insertcols!(df, 1, :id=>["$i" for i in sheetMatrix[2:end, 1]])
        dfArray[i] = df
    end
    return dfArray
end

	
"""
meanOfDfArray(dfArray)
Compute the column means of each DataFrame in dfArray and returns a new DataFrame contaning all computed means
# Examples
```jldoctest
julia> meanOfDfArray(dfArray)
8
```
"""
function meanOfDfArray(dfArray)
    means = Array{Array{Float64,2}, 1}(undef, length(dfArray))
    for i in eachindex(dfArray)
        means[i] = mean(Matrix(dfArray[i][:,2:end]), dims=1)'
    end
    return means
end

dfs = excelSheets2DfArray("country1.xlsx")
res = meanOfDfArray(dfs)
resDf = DataFrame(hcat(res...), [Symbol("mean$i") for i in 1:3])











function data2CSV(filePath)
    file =  XLSX.readxlsx(filePath)
    sheetNames = [i.name for i in file.workbook.sheets]
    avgMatrix = zeros(20, length(sheetNames))

    for (i, sheet) in enumerate(sheetNames)
        sheet = file[sheet]
        sheetMatrix = convert(Array{Float64,2}, sheet[:][2:32,2:21])
        avg = sum(sheetMatrix, dims=1)
        avgMatrix[:,i] = avg
    end
    avgDf = DataFrame(avgMatrix, Symbol.(sheetNames))
    insertcols!(avgDf, 1, :year=>["$i" for i in 1998:2017])
    names!(avgDf, [:年份, :劳动人数, :全时人员当量, :工资, :政府投资, :日常性支出, :论文, :专利])
    CSV.write("$(filePath)Avg.csv", avgDf)
end

fileNames = ["uni", "ins", "firms"]


data2CSV("uni.xlsx")
data2CSV("ins.xlsx")
data2CSV("firms.xlsx")
