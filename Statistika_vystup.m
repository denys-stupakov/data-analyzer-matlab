function Statistika_vystup()
    currentDir = pwd;
    filename = fullfile(currentDir, 'DataInput', 'SVP-Statistika.xlsx');
    inputSheetName = 'VstupneData';
    outputSheetName = 'VystupneData';
    %Clearing Sheet before adding new information
    Table = readtable(filename, 'Sheet', outputSheetName, 'VariableNamingRule', 'preserve');
    [row, col] = find(~ismissing(Table));
    row = max(row) + 1;
    col = max(col) + 1;
    emptyTable = array2table(cell(row, col), 'VariableNames', strcat("Name", string(1:col)));
    writetable(emptyTable, filename, 'Sheet', outputSheetName, 'Range', 'A1', 'WriteVariableNames', false);

    T = readtable(filename, 'Sheet', inputSheetName, 'VariableNamingRule', 'preserve');
    cities_population_under_1000 = T(T.('Počet obyvateľov sídla') < 1000, {'Názov sídla', 'Počet obyvateľov sídla'});
    cities_population_under_1000 = sortrows(cities_population_under_1000, 'Počet obyvateľov sídla', 'descend');
    cities_population_over_50000 = T(T.('Počet obyvateľov sídla') > 50000, {'Názov sídla', 'Počet obyvateľov sídla'});
    cities_population_over_50000 = sortrows(cities_population_over_50000, 'Počet obyvateľov sídla', 'descend');
    
    writetable(cities_population_over_50000, filename, 'Sheet', outputSheetName);
    number_of_rows = height(cities_population_over_50000) + 3;
    needed_cell = ['A', num2str(number_of_rows)];
    writetable(cities_population_under_1000, filename, 'Sheet', outputSheetName, 'WriteRowNames', true, 'Range', needed_cell);

    excel = actxserver('Excel.Application');
    excel.Visible = true;
    workbook = excel.Workbooks.Open(filename);
    sheet = workbook.Sheets.Item(outputSheetName);
    
    range = sheet.UsedRange;
    range.Interior.ColorIndex = -4142;
    range.HorizontalAlignment = 1;
    range.VerticalAlignment = -4160;

    str = sprintf("A1:B%d", height(cities_population_over_50000) + 1);
    range = sheet.Range(str);
    range.Interior.Color = 255 + 223*256 + 186*256^2;
    range.HorizontalAlignment = -4108;
    range.VerticalAlignment = -4108;

    str = sprintf("A%d:B%d", number_of_rows, height(cities_population_under_1000) + number_of_rows);
    range = sheet.Range(str);
    range.Interior.Color = 255 + 223*256 + 186*256^2;
    range.HorizontalAlignment = -4108;
    range.VerticalAlignment = -4108;

    range = sheet.Range('A1:B1');
    range.Interior.Color = 255 + 165*256 + 0*256^2;

    str = sprintf("A%d:B%d", number_of_rows, number_of_rows);
    range = sheet.Range(str);
    range.Interior.Color = 255 + 165*256 + 0*256^2;
    
    workbook.Save();
    workbook.Close();
    excel.Quit();
    excel.delete();
end
