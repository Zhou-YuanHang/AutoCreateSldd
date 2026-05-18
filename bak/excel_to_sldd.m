%% excel_to_sldd — 从 Excel 模板生成 Simulink 对象 M 脚本
%  脚本入口，按 F5 运行。
%  自动查找当前目录下的 .xlsx → 选择要处理的 sheet → 生成：
%    - xxx_objects.m  （Signal / Parameter / Const / Bus / BusElement）
%    - EnumName.m     （每个枚举一个类定义文件）
%
%  所有转换逻辑在 Excel2Workspace 函数中，不污染工作区。

clear; clc;
fprintf('===============================================\n');
fprintf('      Excel 数据字典导入工具\n');
fprintf('===============================================\n\n');

% 查找当前文件夹下的第一个 .xlsx
currDir = pwd;
xlsFiles = dir(fullfile(currDir, '*.xlsx'));
if isempty(xlsFiles)
    xlsFiles = dir(fullfile(currDir, '*.xls'));
end

if ~isempty(xlsFiles)
    defaultFile = xlsFiles(1).name;
    fprintf('找到 Excel 文件: %s\n', defaultFile);
    choice = input('是否使用此文件？(Y/N, 默认 Y): ', 's');
    if isempty(choice) || upper(choice) == 'Y'
        filename = defaultFile;
        filepath = currDir;
    else
        [filename, filepath] = uigetfile( ...
            {'*.xlsx;*.xls', 'Excel 文件 (*.xlsx, *.xls)'}, ...
            '选择数据字典 Excel 文件');
    end
else
    [filename, filepath] = uigetfile( ...
        {'*.xlsx;*.xls', 'Excel 文件 (*.xlsx, *.xls)'}, ...
        '选择数据字典 Excel 文件');
end

if isequal(filename, 0)
    disp('已取消');
    return;
end

fullpath = fullfile(filepath, filename);
fprintf('输入文件: %s\n', fullpath);

% 获取 sheet 列表
[~, AllSheets] = xlsfinfo(fullpath);
fprintf('找到的工作表: %s\n', strjoin(AllSheets, ', '));

% 输出文件名
[~, baseName, ~] = fileparts(filename);
outputFilename = fullfile(filepath, [baseName, '_objects.m']);

% 调用核心函数（所有转换逻辑在里面）
Excel2Workspace(AllSheets, fullpath, outputFilename, filename);

fprintf('\n===== 完成 =====\n');
fprintf('主脚本: %s\n', outputFilename);
fprintf('枚举 .m 文件在同一目录。\n');

% 清理本脚本的临时变量
clear currDir xlsFiles defaultFile choice filename filepath ...
      fullpath AllSheets baseName outputFilename
