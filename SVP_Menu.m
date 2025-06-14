folderName = 'DataInput';
fileName = 'SVP-Statistika.xlsx';
filePath = fullfile(folderName, fileName);

if ~isfile(filePath)
    disp('Súbor SVP-Statistika.xlsx nie je k dispozícii.');
    return;
end

teamInfo = {
    'Name', 'Surname', 'Role';
    'Denys', 'Stupakov', 'Project Manager';
    'Oleksandr', 'Bereza', 'Tester';
    'Renat', 'Moskalenko', 'Data Analyser';
    'Yevhen', 'Kozirovskyi', 'Developer'
};

slovakDescription = {
    '', '', '';
    'Popis aplikácie:', '', '';
    '', 'Naša aplikácia obsahuje hlavné menu, ktoré pozostáva zo štyroch tlačidiel:', '';
    '', '1. Ukonči program - ukončenie aplikácie,', '';
    '', '2. Grafy - otvorenie grafickej kalkulačky,', '';
    '', '3. Matice - výpočet a operácie s maticami,', '';
    '', '4. Štatistika - spracovanie tabuľky Excel a poskytovanie dôležitých informácií späť do súboru Excel.', '';
    '', '', '';
    'Popis zadania:', '', '';
    '', 'Zadanie č. 3 (projekt):', '';
    '', 'Cieľom projektu je vytvoriť aplikáciu v MATLABe, ktorá vykoná funkcionalitu uvedenú v popise aplikácie vyššie.', '';
    '', 'Riešenie vypracováva tím pozostávajúci zo štyroch členov: Project Manager, Developer, Tester a Data Analyser.', '';
    '', 'Úlohy a zodpovednosti:', '';
    '', '- Project Manager: Zodpovedá za komunikáciu s vyučujúcimi, obhajobu riešenia, rozdelenie úloh a asistenciu ostatným členom tímu.', '';
    '', '- Developer: Vytvára samotný program vrátane grafického a komunikačného rozhrania.', '';
    '', '- Tester: Testuje funkcionalitu programu na vhodných a nevhodných vstupoch a kontroluje jeho použiteľnosť.', '';
    '', '- Data Analyser: Zabezpečuje získavanie a spracovanie potrebných údajov.', ''
};

try
    writetable(table(), filePath, 'Sheet', 'ZakladneInfo');
catch ME
    disp('Došlo k chybe pri vymazávaní obsahu hárka ZakladneInfo.');
    disp(ME.message);
    return;
end

outputData = [teamInfo; slovakDescription];

try
    writecell(outputData, filePath, 'Sheet', 'ZakladneInfo');
    disp('Informácie o tíme a popis boli úspešne zapísané do hárku ZakladneInfo.');
catch ME
    disp('Došlo k chybe pri zapisovaní do súboru.');
    disp(ME.message);
end
menu();