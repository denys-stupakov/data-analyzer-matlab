function graf_kumulat_fi(x, y)
    fig = figure('Visible', 'off');
    plot(x, y, '-o');
    xlabel('i');
    ylabel('Fi');
    title('Kumulativna relativna početnost pre nadmorskú výšku');

    outputPath = fullfile(pwd, 'Grafs', 'graf_kumulat_fi.png');
    if ~isfolder(fullfile(pwd, 'Grafs'))
        mkdir(fullfile(pwd, 'Grafs'));
    end
    saveas(fig, outputPath);
    close(fig);
end