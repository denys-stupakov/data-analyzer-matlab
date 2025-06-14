function add_graphs_side_by_side()
    graphFolder = fullfile(pwd, 'Grafs');
    
    graphFiles = {'graf_fi.png', 'graf_kumulat_fi.png', 'graf_kumulat_ni.png', 'graf_ni.png'};
  
    excelFile = fullfile(pwd, 'DataInput', 'SVP-Statistika.xlsx');
    sheetName = 'PomocneUdaje';

    if ~isfile(excelFile)
        error('Excel file does not exist: %s', excelFile);
    end

    excel = actxserver('Excel.Application');
    excel.Visible = true;
    workbook = excel.Workbooks.Open(excelFile);
    
    try
        try
            sheet = workbook.Sheets.Item(sheetName);
        catch
            sheet = workbook.Sheets.Add([], workbook.Sheets.Item(workbook.Sheets.Count));
            sheet.Name = sheetName;
        end

        shapes = sheet.Shapes;
        for i = shapes.Count:-1:1
            shapes.Item(i).Delete();
        end
        disp('Deleted all existing shapes on the sheet.');

        startCell = 'D1';
        horizontalSpacing = 400;
        verticalSpacing = 0;

        range = sheet.Range(startCell);
        initialLeft = range.Left;
        initialTop = range.Top;

        left = initialLeft;
        top = initialTop;

        for i = 1:length(graphFiles)
            graphPath = fullfile(graphFolder, graphFiles{i});
            if ~isfile(graphPath)
                warning('Graph file does not exist: %s', graphPath);
                continue;
            end

            sheet.Shapes.AddPicture(graphPath, 0, 1, left, top, -1, -1);
            disp(['Added graph: ', graphFiles{i}, ' at position Left = ', num2str(left), ', Top = ', num2str(top)]);

            left = left + horizontalSpacing;

            if i > 1 && mod(i, 4) == 0
                left = initialLeft;
                top = top + verticalSpacing;
            end
        end

        workbook.Save();
    catch ME
        disp('Error while interacting with Excel:');
        disp(ME.message);
    end

    try
        workbook.Close();
        excel.Quit();
        delete(excel);
    catch ME
        disp('Error while closing Excel:');
        disp(ME.message);
    end
end