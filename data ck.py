import pandas

dataFrameck = pandas.read_excel("2025 Lytle Density V4.xlsx", header=1)

dictionaryck = {}

for i in range(len(dataFrameck)):
    blockValue = dataFrameck.loc[i, "BLOCK"]
    plotValue = dataFrameck.loc[i, "PLOT"]
    speciesValue = dataFrameck.loc[i, "SPECIES"]
    totalValue = dataFrameck.loc[i, 1:12].sum()

    if blockValue not in dictionaryck:
        dictionaryck[blockValue] = {
            "BN": {},
            "BS": {},
            "UN": {},
            "US": {}
        }

    if speciesValue not in dictionaryck[blockValue][plotValue]:
        dictionaryck[blockValue][plotValue][speciesValue] = 0

    dictionaryck[blockValue][plotValue][speciesValue] += totalValue #end of part 1


plotDescriptions = {
    "BN": ["Burned", "Excluded"],
    "BS": ["Burned", "Present"],
    "UN": ["Unburned", "Excluded"],
    "US": ["Unburned", "Present"]
}

allSpecies = set()

for blockValue in dictionaryck:
    for plotValue in dictionaryck[blockValue]:
        for speciesValue in dictionaryck[blockValue][plotValue]:
            allSpecies.add(speciesValue)

allSpecies = sorted(allSpecies)

rowsForNewSheet = []

for plotValue in ["BN", "BS", "UN", "US"]:
    for blockValue in [1, 2, 3, 4, 5]:
        rowDictionary = {
            "DATE": dataFrameck.loc[0, "DATE"],
            "BLOCK": blockValue,
            "PLOT": plotValue,
            "FIRE": plotDescriptions[plotValue][0],
            "RODENTS": plotDescriptions[plotValue][1]
        }

        for speciesValue in allSpecies:
            rowDictionary[speciesValue] = 0

        if blockValue in dictionaryck and plotValue in dictionaryck[blockValue]:
            for speciesValue in dictionaryck[blockValue][plotValue]:
                rowDictionary[speciesValue] = dictionaryck[blockValue][plotValue][speciesValue]

        rowsForNewSheet.append(rowDictionary)

newSheetDataFrameck = pandas.DataFrame(rowsForNewSheet)

fixedColumns = ["DATE", "BLOCK", "PLOT", "FIRE", "RODENTS"]
speciesColumns = allSpecies

newSheetDataFrameck = newSheetDataFrameck[fixedColumns + speciesColumns]

print(newSheetDataFrameck)

with pandas.ExcelWriter(
    "2025 Lytle Density V4.xlsx",
    engine="openpyxl",
    mode="a",
    if_sheet_exists="replace"
) as writer:
    newSheetDataFrameck.to_excel(writer, sheet_name="Species Totals Wide", index=False)

 
#print(dictionaryck[1]["BS"]["ERCI"])