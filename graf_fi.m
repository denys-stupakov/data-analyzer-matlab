function graf_fi(x, y)
    fig = figure('Visible', 'off');
    pie(y);
    xlabel('i');
    ylabel('fi');
    title('Relativna početnosť pre nadmorskú výšku');
    
    values = arrayfun(@(v) num2str(v), x, 'UniformOutput', false);
    legend(values, 'Location', 'best', 'Orientation', 'horizontal');
    
    outputPath = fullfile(pwd, 'Grafs', 'graf_fi.png');
    if ~isfolder(fullfile(pwd, 'Grafs'))
        mkdir(fullfile(pwd, 'Grafs'));
    end
    saveas(fig, outputPath);
    close(fig);
end