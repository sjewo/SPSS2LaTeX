# -*- coding: utf8 -*-
# version 0.0.1
# author: Sebastian Jeworutzki
# contact: Sebastian.Jeworutzki@ruhr-uni-bochum.de

import SpssClient
SpssClient.StartClient()

OutputDoc = SpssClient.GetDesignatedOutputDoc()
OutputItems = OutputDoc.GetOutputItems()

env = u"" # Enthält Tabelledefinition
head = u"" # Enthält Tabellenüberschrift
tab =  u"" # Enthält die Ausgabetabelle

for index in range(OutputItems.Size()):
    OutputItem = OutputItems.GetItemAt(index)
    if OutputItem.GetType() == SpssClient.OutputItemType.HEAD:
        lastheader = index

OutputItem = OutputDoc.GetCurrentItem()
if OutputItem.GetType() == SpssClient.OutputItemType.PIVOT:
    PivotTable = OutputItem.GetSpecificType()
    ColumnLabel = PivotTable.ColumnLabelArray()
    RowLabel = PivotTable.RowLabelArray()

    # Tabular einleiten            
    env = env + "\\begin{tabular}{"
    for i in range(RowLabel.GetNumColumns()):
        env = env + "l"
    for i in range(DataCells.GetNumColumns()):
        env = env + "r"
    env = env + "}\n"
    
    # Tabellenkopf setzten
    for i in range(ColumnLabel.GetNumRows()):
        for k in range(RowLabel.GetNumColumns()):
            head = head + " & "
        for j in range(ColumnLabel.GetNumColumns()):
            if j==ColumnLabel.GetNumColumns(): # Bis auf die letzte Zelle ein & anhängen
                head = head + ColumnLabel.GetValueAt(i,j)
            else:
                head = head + ColumnLabel.GetValueAt(i,j) + " & "
        head = head + " \\\\ \n"
    head = head + "\\hrule  \n" # Linie unter der Überschrift

    # Tabelleninhalt setzten
    DataCells = PivotTable.DataCellArray()
    for i in range(DataCells.GetNumRows()):
        for k in range(RowLabel.GetNumColumns()):
            tab = tab + RowLabel.GetValueAt(i,k) + " & "
        for j in range(DataCells.GetNumColumns()):
            if j==DataCells.GetNumColumns(): # Bis auf die letzte Zelle ein & anhängen
                tab = tab + DataCells.GetValueAt(i,j)
            else:
                tab = tab + DataCells.GetValueAt(i,j) + " & "
        tab = tab + " \\\\ \n"

out = env + head + tab + u"\\end{tabular}"
print(out)

# Get the root header item
root = OutputItems.GetItemAt(lastheader).GetSpecificType()
# Create a new text item
newText = OutputDoc.CreateTextItem(out)
# Append the new text item to the new header item
root.InsertChildItem(newText,root.GetChildCount())


SpssClient.StopClient()
