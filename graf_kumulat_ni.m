function graf_kumulat_ni(x, y)
    % Generate and save the plot
    fig = figure('Visible', 'off');
    plot(x, y, '-o');
    xlabel('i');
    ylabel('Ni');
    title('Kumulativna absolutna početnost pre nadmorskú výšku');

    outputPath = fullfile(pwd, 'Grafs', 'graf_kumulat_ni.png');
    if ~isfolder(fullfile(pwd, 'Grafs'))
        mkdir(fullfile(pwd, 'Grafs'));
    end
    saveas(fig, outputPath);
    close(fig);
end
