function test(Mesiac, Nadmorska_vyska)
    currentDir = pwd;
    filename = fullfile(currentDir, 'DataInput', 'SVP-Statistika.xlsx');
    inputSheetName = 'VstupneData';
    outputSheetName = 'VystupneData';
    
    Input_Table = readtable(filename, 'Sheet', inputSheetName, 'VariableNamingRule', 'preserve');
    switch Mesiac
        case 'Februar'
            month = 2;
        case 'Marec'
            month = 3;
        case 'April'
            month = 4;
        case 'Maj'
            month = 5;
        case 'Jun'
            month = 6;
        case 'Jul'
            month = 7;
        case 'August'
            month = 8;
        case 'September'
            month = 9;
        case 'Oktober'
            month = 10;
        case 'November'
            month = 11;
        case 'December'
            month = 12;
        otherwise
            month = 1;
    end

    v_vyska = Input_Table(Input_Table.('Nadmorská výška (m)') > Nadmorska_vyska, {'Priemerná mesačná teplota (°C)', 'Priemerný mesačný úhrn zrážok (mm)'});
    m_vyska = Input_Table(Input_Table.('Nadmorská výška (m)') < Nadmorska_vyska, {'Priemerná mesačná teplota (°C)', 'Priemerný mesačný úhrn zrážok (mm)'});

    Output_Sheet = readtable(filename, 'Sheet', outputSheetName, 'VariableNamingRule', 'preserve');
    [row, col] = find(~ismissing(Output_Sheet));
    row = max(row) + 3;
    needed_cell = ['A', num2str(row)];

    header1 = "Priemerna mesačna teplota a priemerný mesačný úhrn zrážok v mestách s nadmorskou výškou vyššou ako " + num2str(Nadmorska_vyska) + " m v mesiaci " + Mesiac;
    Header1 = table(header1, 'VariableNames', {'Name'});
    header2 = "Priemerna mesačna teplota a priemerný mesačný úhrn zrážok v mestách s nadmorskou výškou menšou ako " + num2str(Nadmorska_vyska) + " m v mesiaci " + Mesiac;
    Header2 = table(header2, 'VariableNames', {'Name'});

    if(~isempty(v_vyska))
        h = height(v_vyska);
        avg_vt = 0;
        avg_vu = 0;
        for i = 1:h
            temp_row = v_vyska .('Priemerná mesačná teplota (°C)'){i};
            temp_array = str2double(split(temp_row, ','));
            avg_vt = avg_vt + temp_array(month);
            temp_row = v_vyska .('Priemerný mesačný úhrn zrážok (mm)'){i};
            temp_array = str2double(split(temp_row, ','));
            avg_vu = avg_vu + temp_array(month);
        end
        avg_vt = avg_vt/h;
        avg_vt = round(avg_vt, 3);
        avg_vu = avg_vu/h;
        avg_vu = round(avg_vu, 3);

        writetable(Header1, filename, 'Sheet', outputSheetName, 'WriteVariableNames', false, 'Range', needed_cell);
        needed_cell1 = ['A', num2str(row + 1)];
        dataTable = table(avg_vt, avg_vu, 'VariableNames', {'Priemerná teplota (°C)', 'Priemerný úhrn zrážok (mm)'});
        writetable(dataTable, filename, 'Sheet', outputSheetName, 'WriteRowNames', true, 'Range', needed_cell1);

        excel = actxserver('Excel.Application');
        excel.Visible = true;
        workbook = excel.Workbooks.Open(filename);
        sheet = workbook.Sheets.Item(outputSheetName);
        
        str = sprintf("A%d:B%d", row + 1, row + 2);
        range = sheet.Range(str);
        range.HorizontalAlignment = -4108;
        range.VerticalAlignment = -4108;
        range.Interior.Color = 210 + 232*256 + 255*256^2;

        str = sprintf("A%d:B%d", row + 1, row + 1);
        range = sheet.Range(str);
        range.Interior.Color = 135 + 206*256 + 235*256^2;

        workbook.Save();
        workbook.Close();
        excel.Quit();
        excel.delete();
    end
    if(~isempty(m_vyska))
        avg_mt = 0;
        avg_mu = 0;
        h = height(m_vyska);
        for i = 1:h
            temp_row = m_vyska .('Priemerná mesačná teplota (°C)'){i};
            temp_array = str2double(split(temp_row, ','));
            avg_mt = avg_mt + temp_array(month);
            temp_row = m_vyska .('Priemerný mesačný úhrn zrážok (mm)'){i};
            temp_array = str2double(split(temp_row, ','));
            avg_mu = avg_mu + temp_array(month);
        end
        avg_mt = avg_mt/h;
        avg_mt = round(avg_mt, 3);
        avg_mu = avg_mu/h;
        avg_mu = round(avg_mu, 3);

        dataTable = table(avg_mt, avg_mu, 'VariableNames', {'Priemerná teplota (°C)', 'Priemerný úhrn zrážok (mm)'});
        if(~isempty(v_vyska))
            needed_cell = ['G', num2str(row)];
            needed_cell1 = ['G', num2str(row + 1)];
            d_c = sprintf("G%d:H%d", row + 1, row + 2);
            h_c = sprintf("G%d:H%d", row + 1, row + 1);
        else
            needed_cell = ['A', num2str(row)];
            needed_cell1 = ['A', num2str(row + 1)];
            d_c = sprintf("A%d:B%d", row + 1, row + 2);
            h_c = sprintf("A%d:B%d", row + 1, row + 1);
        end
        writetable(Header2, filename, 'Sheet', outputSheetName, 'WriteVariableNames', false, 'Range', needed_cell);
        writetable(dataTable, filename, 'Sheet', outputSheetName, 'WriteRowNames', true, 'Range', needed_cell1);

        excel = actxserver('Excel.Application');
        excel.Visible = true;
        workbook = excel.Workbooks.Open(filename);
        sheet = workbook.Sheets.Item(outputSheetName);
        
        range = sheet.Range(d_c);
        range.HorizontalAlignment = -4108;
        range.VerticalAlignment = -4108;
        range.Interior.Color = 210 + 232*256 + 255*256^2;

        range = sheet.Range(h_c);
        range.Interior.Color = 135 + 206*256 + 235*256^2;
        
        workbook.Save();
        workbook.Close();
        excel.Quit();
        excel.delete();
    end
end