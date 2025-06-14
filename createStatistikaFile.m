function createStatistikaFile()
    filePath = fullfile('DataInput', 'SVP-Statistika.xlsx');
    
    if ~isfolder('DataInput')
        mkdir('DataInput');
    end
    
    sheetNames = {'ZakladneInfo', 'VstupneData', 'VystupneData', 'Charakteristiky', 'PomocneUdaje'};
    
    for i = 1:length(sheetNames)
        writecell({'Placeholder'}, filePath, 'Sheet', sheetNames{i});
    end
end