function graf_ni(x, y)
    fig = figure('Visible', 'off');
    bar(x, y);
    xlabel('i');
    ylabel('ni');
    title('Absolutna početnost pre nadmorsku vysku');
    
    outputPath = fullfile(pwd, 'Grafs', 'graf_ni.png');
    if ~isfolder(fullfile(pwd, 'Grafs'))
        mkdir(fullfile(pwd, 'Grafs'));
    end
    saveas(fig, outputPath);
    close(fig);
end