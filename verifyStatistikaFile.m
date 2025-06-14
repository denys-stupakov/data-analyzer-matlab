function exists = verifyStatistikaFile()
    filePath = fullfile('DataInput', 'SVP-Statistika.xlsx');
    exists = isfile(filePath);
end