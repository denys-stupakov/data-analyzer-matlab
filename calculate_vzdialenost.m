function calculate_vzdialenost(distance)
    currentDir = pwd;
    filename = fullfile(currentDir, 'DataInput', 'SVP-Statistika.xlsx');
    inputSheetName = 'VstupneData';
    outputSheetName = 'VystupneData';

    Input_Table = readtable(filename, 'Sheet', inputSheetName, 'VariableNamingRule', 'preserve');
    Output_Table = Input_Table(Input_Table.('Najkratšia cestná vzdialenosť od Košíc (v km)') < distance & Input_Table.('Najkratšia cestná vzdialenosť od Košíc (v km)') ~= 0, {'Názov sídla','Najkratšia cestná vzdialenosť od Košíc (v km)'});
    if(height(Output_Table) > 0)
        Output_Table = sortrows(Output_Table, 'Najkratšia cestná vzdialenosť od Košíc (v km)', 'ascend');
        Output_Table.Properties.VariableNames{'Najkratšia cestná vzdialenosť od Košíc (v km)'} = 'Vzdialenost''';
    
        Output_Sheet = readtable(filename, 'Sheet', outputSheetName, 'VariableNamingRule', 'preserve');
        [row, col] = find(~ismissing(Output_Sheet));
        row = max(row) + 3;
        needed_cell = ['A', num2str(row)];
    
        header = "Všetky sídla vzdialenost' ktorých je menšia ako " + num2str(distance) + " km od Košic";
        Header = table(header, 'VariableNames', {'Name'});
        writetable(Header, filename, 'Sheet', outputSheetName, 'WriteVariableNames', false, 'Range', needed_cell);
        needed_cell = ['A', num2str(row + 1)];
        writetable(Output_Table, filename, 'Sheet', outputSheetName, 'WriteRowNames', true, 'Range', needed_cell);
        
        excel = actxserver('Excel.Application');
        excel.Visible = true;
        workbook = excel.Workbooks.Open(filename);
        sheet = workbook.Sheets.Item(outputSheetName);
       
        h = height(Output_Table);
        str = sprintf("A%d:B%d", row + 1, h + row + 1);
        range = sheet.Range(str);
        range.HorizontalAlignment = -4108;
        range.VerticalAlignment = -4108;

        str = sprintf("A%d:B%d", row + 1, row + 1);
        range = sheet.Range(str);
        range.Interior.Color = 152 + 251*256 + 152*256^2;

        str = sprintf("A%d:B%d", row + 2, h + row + 1);
        range = sheet.Range(str);
        range.Interior.Color = 230 + 255*256 + 230*256^2;

        workbook.Save();
        workbook.Close();
        excel.Quit();
        excel.delete();
    end
end