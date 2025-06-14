function MaticeVysledky()
    fileID = fopen('DataInput/Matice.txt', 'r');
    
    line = fgetl(fileID);
    matrixA = [];
    matrixB = [];
    currentMatrix = 0;
    
    while ischar(line)
        if contains(line, 'Matica A')
            currentMatrix = 1;
        elseif contains(line, 'Matica B')
            currentMatrix = 2;
        else
            if ~isempty(strtrim(line))
                numbers = str2num(line);
                if currentMatrix == 1
                    matrixA = [matrixA; numbers];
                elseif currentMatrix == 2
                    matrixB = [matrixB; numbers];
                end
            end
        end
        line = fgetl(fileID);
    end
    fclose(fileID);

    r = rank(matrixB);
    d = det(matrixB);
    i = inv(matrixB);
    savepath = "DataOutput";
    if ~exist(savepath, 'dir')
        mkdir(savepath);
    end
    filename = fullfile(savepath, 'MaticaVysledky.txt');
    fp = fopen(filename,'w');
    fprintf(fp, 'Hodnosť matice B: %d\nDeterminant matice B: %d\n', r,d);
    fprintf(fp, '\nInverzná matica ku matice (B):\n');
    for row = 1:size(i, 1)
        fprintf(fp, '%d\t', i(row, :));
        fprintf(fp, '\n');
    end
    fclose(fp);
end