
function GenerateMatrices()
    m = randi([2, 25]);
    n = randi([2, 25]);
    a = -100;
    b = 100;
    A = randi([a, b], [m, n]);
    B = A * A';

    Matice = {A, B};

    savepath = 'DataInput';
    if ~exist(savepath, 'dir')
        mkdir(savepath);
    end

    filename = fullfile(savepath, 'Matice.txt');
    fp = fopen(filename, 'wt');
    if fp == -1
        error('Nie je možné otvoriť súbor na zápis.');
    end

    for i = 1:numel(Matice)
        fprintf(fp, 'Prvky matice %c su z intervalu [%d, %d].\n', 64+i, a, b);

        fprintf(fp, 'Matica %c:\n', 64+i);
        mat = Matice{i};
        
        for row = 1:size(mat, 1)
            fprintf(fp, '%d\t', mat(row, :));
            fprintf(fp, '\n');
        end
        fprintf(fp, '\n');
    end
    
    fclose(fp);
end