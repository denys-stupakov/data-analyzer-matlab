function function_Statistika()
    currentDir = pwd;
    filename = fullfile(currentDir, 'DataInput', 'SVP-Statistika.xlsx');
    sheetname = 'VstupneData';

    excel = actxserver('Excel.Application');
    excel.Visible = true;
    workbook = excel.Workbooks.Open(filename);
    sheet = workbook.Sheets.Item('Charakteristiky');
    
    range = sheet.UsedRange;
    range.Interior.ColorIndex = -4142;
    range.HorizontalAlignment = 1;
    range.VerticalAlignment = -4160;

    workbook.Save();
    workbook.Close();
    excel.Quit();
    excel.delete();
    
    %Clearing Sheet before adding new information
    emptyTable = array2table(cell(150, 40), 'VariableNames', strcat("Name", string(1:40)));
    writetable(emptyTable, filename, 'Sheet', 'Charakteristiky', 'Range', 'A1', 'WriteVariableNames', false);



    Tin = readtable(filename, 'Sheet', sheetname, 'VariableNamingRule', 'preserve');
    Tout = table();
    [n, x] = size(Tin);

    m = ceil(sqrt(n));
    if mod(m, 2) ~= 0
        m = m + 1;
    end

    xmin = min(Tin{:, 8});
    xmax = max(Tin{:, 8});
    R = xmax - xmin;
    h = ceil(R/m);

    Tout(:, 1) = table((1:m)'); 
    Tout.Properties.VariableNames{1} = 'i';

    Tout(:, 2) = table((xmin-1:h:xmin + (m - 1) * h)');
    Tout.Properties.VariableNames{2} = 'ai';

    Tout(:, 3) = table(Tout{:, 2} + h);
    Tout.Properties.VariableNames{3} = 'bi';

    Tout(:, 4) = table((Tout{:, 2} + Tout{:, 3})/2);
    Tout.Properties.VariableNames{4} = 'xi';
    temp = Tin{:, 8};
    Ni = {};
    ni = {};
    for i = 1:height(Tout)
        Ni{i} = temp(temp < Tout{i, 3});
        Ni{i} = size(Ni{i}, 1);
        if i >= 2
            ni{i} = Ni{i} - Ni{i - 1};
        elseif i == 1
            ni{1} = Ni{1};
        end
    end

    Tout(:, 5) = table(cell2mat(ni'));
    Tout.Properties.VariableNames{5} = 'ni';

    graf_ni(1:m, Tout{:, 5});

    Tout(:, 6) = table(cell2mat(Ni'));
    Tout.Properties.VariableNames{6} = 'Ni';
    graf_kumulat_ni(1:m, Tout{:, 6});

    Tout(:, 7) = table(Tout{:, 5} / n);
    Tout.Properties.VariableNames{7} = 'fi';

    graf_fi(1:m, Tout{:, 7});

    Tout(:, 8) = table(Tout{:, 6} / n);
    Tout.Properties.VariableNames{8} = 'Fi';

    graf_kumulat_fi(1:m, Tout{:, 8});
    add_graphs_to_excel();

    Tout(:, 9) = table(Tout{:, 4} .* Tout{:, 5});
    Tout.Properties.VariableNames{9} = 'xi * ni';
    AVG = sum(Tout{:, 9}) / n;

    Tout(:, 10) = table(Tout{:, 5} .* ((Tout{:, 4} - AVG) .^2) );
    Tout.Properties.VariableNames{10} = 'ni * (xi - AVG)^2';

    [maxValue, maxIndex] = max(Tout{:, 5});

    if maxIndex == 1
        a0 = Tout{maxIndex, 2};
        d1 = maxValue;          
        d2 = maxValue - Tout{maxIndex + 1, 5};
        MODUS = a0 + h * d1 / (d1 + d2);
    elseif maxIndex == height(Tout)
        a0 = Tout{maxIndex, 2}; 
        d1 = maxValue - Tout{maxIndex - 1, 5}; 
        d2 = maxValue;          
        MODUS = a0 + h * d1 / (d1 + d2);
    else
        a0 = Tout{maxIndex, 2}; 
        d1 = maxValue - Tout{maxIndex - 1, 5}; 
        d2 = maxValue - Tout{maxIndex + 1, 5}; 
        MODUS = a0 + h * d1 / (d1 + d2);
    end

    medianCol = Tout{:, 8};
    medianValue = min(medianCol(medianCol >= 0.5));
    medianIndex = find(medianCol == medianValue, 1);
    ae = Tout{medianIndex, 2};
    ne = Tout{medianIndex, 5};
    if medianIndex > 1
        Ne = Tout{medianIndex - 1, 6};
    else
        Ne = 0;
    end
    MEDIAN = ae + h * (((n + 1) / 2 - Ne) / ne);
    ROZPTYL = sum(Tout{:, end}) / (n - 1);
    SmerodajnaOdchylka = sqrt(ROZPTYL);

    h = height(Tout) + 3;

    params = {'n'; 'm'; 'Xmin'; 'Xmax'; 'R'; 'h'};
    values = {n; m; xmin; xmax; R; h};
    Param1Table = table(params, values, 'VariableNames', {'Name1', 'Name2'});

    %modus
    params = {'Modus'; 'a0'; 'd1'; 'd2'; ' '; 'x'};
    values = {' '; a0; d1; d2; ' '; MODUS};
    Param2table = table(params, values, 'VariableNames', {'Name1', 'Name2'});

    %median
    params = {'Median'; 'ae'; 'ne'; 'Ne'; ' '; 'x'};
    values = {' '; ae; ne; Ne; ' '; MEDIAN};
    Param3table = table(params, values, 'VariableNames', {'Name1', 'Name2'});

    Param4table = table(ROZPTYL, 'VariableNames', {'Rozptyl  '});
    Param5table = table(SmerodajnaOdchylka, 'VariableNames', {'Smerodajna odchylka'});
    Param6table = table(AVG, 'VariableNames', {'Aritmeticky priemer'});

    writetable(Param3table, filename, 'Sheet', 'Charakteristiky', 'Range', ['H', num2str(h)], 'WriteVariableNames', false);
    writetable(Param4table, filename, 'Sheet', 'Charakteristiky', 'Range', ['K', num2str(h)]);
    writetable(Param5table, filename, 'Sheet', 'Charakteristiky', 'Range', ['M', num2str(h)]);
    writetable(Param6table, filename, 'Sheet', 'Charakteristiky', 'Range', ['O', num2str(h)]);
    writetable(Tout, filename, 'Sheet', 'Charakteristiky', 'Range', 'A1');
    writetable(Param1Table, filename, 'Sheet', 'Charakteristiky', 'Range', ['B', num2str(h)], 'WriteVariableNames', false);
    writetable(Param2table, filename, 'Sheet', 'Charakteristiky', 'Range', ['E', num2str(h)], 'WriteVariableNames', false);

    excel = actxserver('Excel.Application');
    excel.Visible = true;
    workbook = excel.Workbooks.Open(filename);
    sheet = workbook.Sheets.Item('Charakteristiky');

    range = sheet.Range('A1:J1');
    range.Interior.Color = 255*255^2 + 171*255 + 61;
    str = sprintf("A2:A%d", m + 1);
    range = sheet.Range(str);
    range.Interior.Color = 255*255^2 + 179*255 + 112;
    str = sprintf("B2:J%d", m + 1);
    range = sheet.Range(str);
    range.Interior.Color = 255*255^2 + 238*255 + 219;
    str = sprintf("B%d:C%d", m + 3, m + 8);
    range = sheet.Range(str);
    range.Interior.Color = 255*255^2 + 238*255 + 219;
    str = sprintf("E%d:F%d", m + 3, m + 8);
    range = sheet.Range(str);
    range.Interior.Color = 255*255^2 + 238*255 + 219;
    str = sprintf("H%d:I%d", m + 3, m + 8);
    range = sheet.Range(str);
    range.Interior.Color = 255*255^2 + 238*255 + 219;
    str = sprintf("K%d:K%d", m + 3, m + 4);
    range = sheet.Range(str);
    range.Interior.Color = 255*255^2 + 238*255 + 219;
    str = sprintf("M%d:M%d", m + 3, m + 4);
    range = sheet.Range(str);
    range.Interior.Color = 255*255^2 + 238*255 + 219;
    str = sprintf("O%d:O%d", m + 3, m + 4);
    range = sheet.Range(str);
    range.Interior.Color = 255*255^2 + 238*255 + 219;

    str = sprintf("A1:J%d", m + 1);
    range = sheet.Range(str);
    range.HorizontalAlignment = -4108;
    range.VerticalAlignment = -4108;

    str = sprintf("B%d:C%d", h, h + 5);
    range = sheet.Range(str);
    range.HorizontalAlignment = -4108;
    range.VerticalAlignment = -4108;

    str = sprintf("E%d:F%d", h, h + 5);
    range = sheet.Range(str);
    range.HorizontalAlignment = -4108;
    range.VerticalAlignment = -4108;

    str = sprintf("H%d:I%d", h, h + 5);
    range = sheet.Range(str);
    range.HorizontalAlignment = -4108;
    range.VerticalAlignment = -4108;

    str = sprintf("K%d:K%d", h, h + 1);
    range = sheet.Range(str);
    range.HorizontalAlignment = -4108;
    range.VerticalAlignment = -4108;

    str = sprintf("M%d:M%d", h, h + 1);
    range = sheet.Range(str);
    range.HorizontalAlignment = -4108;
    range.VerticalAlignment = -4108;

    str = sprintf("O%d:O%d", h, h + 1);
    range = sheet.Range(str);
    range.HorizontalAlignment = -4108;
    range.VerticalAlignment = -4108;

    workbook.Save();
    workbook.Close();
    excel.Quit();
    excel.delete();
end
