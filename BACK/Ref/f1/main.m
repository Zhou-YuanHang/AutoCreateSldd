function main()

clear;clc;
fprintf('===============================================\n');
fprintf('      Excel数据字典导入工具\n');
fprintf('===============================================\n\n');

% 获取当前文件夹路径
currentFolder = pwd;

% 查找当前文件夹下的所有Excel文件
excelFiles = dir(fullfile(currentFolder, '*.xlsx'));
if isempty(excelFiles)
    excelFiles = dir(fullfile(currentFolder, '*.xls'));
end

% 如果找到Excel文件，使用第一个文件
if ~isempty(excelFiles)
    defaultFile = excelFiles(1).name;
    filepath = currentFolder;
    filename = defaultFile;
    fprintf('找到Excel文件: %s\n', filename);
    
    % 询问用户是否使用默认文件
    choice = input('是否使用此文件？(Y/N, 默认Y): ', 's');
    if isempty(choice) || upper(choice) == 'Y'
        fprintf('使用默认文件: %s\n\n', filename);
    else
        % 让用户手动选择文件
        [filename, filepath] = uigetfile({'*.xlsx;*.xls', 'Excel文件 (*.xlsx, *.xls)'},...
            '选择数据字典Excel文件', 'MCU_Template.xlsx');
    end
else
    % 没有找到Excel文件，让用户手动选择
    [filename, filepath] = uigetfile({'*.xlsx;*.xls', 'Excel文件 (*.xlsx, *.xls)'},...
        '选择数据字典Excel文件', 'MCU_Template.xlsx');
end

if isequal(filename, 0)
    disp("用户取消选择excel文件");
    return;
end

% 构建完整文件路径
fullpath = fullfile(filepath, filename);

if ~exist(fullpath, "file")
    error("错误：文件不存在或无法访问。请检查文件路径和权限");
else
    fprintf("文件存在\n");
end

[fid, message] = fopen(fullpath, 'r');
if fid == -1
    error("错误：无法读取文件。原因：%s", message);
else
    fclose(fid);
    fprintf("文件可读\n");
end

% 获取Excel文件中的工作表信息
try
    AllSheets = sheetnames(fullpath);
    if isempty(AllSheets)
        error('无法读取Excel文件，请确保文件格式正确且未被损坏。');
    end
    fprintf('找到的工作表: %s\n', strjoin(AllSheets, ', '));
catch ME
    fprintf('读取Excel文件信息时出错:\n');
    fprintf('  错误信息: %s\n', ME.message);
    return;
end

% 创建输出.m文件
if contains(filename, '.xlsx')
    baseName = strrep(filename, '.xlsx', '');
    outputFilename = [baseName, '_objects.m'];
elseif contains(filename, '.xls')
    baseName = strrep(filename, '.xls', '');
    outputFilename = [baseName, '_objects.m'];
else
    baseName = filename;
    outputFilename = [filename, '_objects.m'];
end

% 完整输出文件路径
fullOutputPath = fullfile(filepath, outputFilename);
fprintf('将生成的MATLAB脚本: %s\n', fullOutputPath);

Excel2Workspace(AllSheets, fullpath, outputFilename, filename);

% 检查生成的脚本文件
if exist(outputFilename, 'file')
    fprintf('已成功生成脚本文件: %s\n', outputFilename);
    scriptContent = fileread(outputFilename);
    evalin('base', scriptContent);  % 在基础工作区执行
    fprintf('脚本已成功执行到基础工作区\n');
else
    error('生成的脚本文件不存在: %s', outputFilename);
end

% 创建数据字典文件名
slddFilename = [baseName, '.sldd'];
fullSLDDPath = fullfile(filepath, slddFilename);
fprintf('将创建的数据字典: %s\n', fullSLDDPath);

% 保存到数据字典
saveAllVarsToDataDictionary(slddFilename);

fprintf('\n===============================================\n');
fprintf('          处理完成！\n');
fprintf('===============================================\n');
fprintf('生成的文件：\n');
fprintf('1. MATLAB脚本: %s\n', fullOutputPath);
fprintf('2. 数据字典: %s\n', fullSLDDPath);
fprintf('\n您可以在以下位置找到这些文件：\n');
fprintf('%s\n', filepath);

% 询问用户是否要打开文件夹
% choice = input('\n是否要打开文件所在文件夹？(Y/N, 默认N): ', 's');
% if ~isempty(choice) && upper(choice) == 'Y'
%     if ispc
%         % Windows系统
%         winopen(filepath);
%     elseif ismac
%         % Mac系统
%         system(['open "' filepath '"']);
%     else
%         % Linux系统
%         system(['xdg-open "' filepath '"']);
%     end
%     fprintf('已打开文件夹\n');
% end

% 询问用户是否要打开数据字典
% choice = input('\n是否要打开数据字典？(Y/N, 默认Y): ', 's');
% if isempty(choice) || upper(choice) == 'Y'
%     try
%         if exist(fullSLDDPath, 'file')
%             dictObj = Simulink.data.dictionary.open(fullSLDDPath);
%             show(dictObj);
%             fprintf('数据字典已打开\n');
%         else
%             fprintf('数据字典文件不存在: %s\n', fullSLDDPath);
%         end
%     catch ME
%         fprintf('打开数据字典时出错: %s\n', ME.message);
%     end
% end

fprintf('\n程序执行完毕！\n');

end

function saveAllVarsToDataDictionary(slddFile)
% SAVEALLVARSTODATADICTIONARY 将基础工作区所有变量保存到Simulink数据字典
%
% 输入参数：
%   slddFile - 数据字典文件名（.sldd）
%
% 输出：
%   显示详细的保存位置信息

    % 获取完整路径
    currentFolder = pwd;
    [filePath, fileName, fileExt] = fileparts(slddFile);
    
    % 如果指定了路径，使用指定路径；否则使用当前文件夹
    if ~isempty(filePath)
        slddFullPath = slddFile;
    else
        slddFullPath = fullfile(currentFolder, slddFile);
    end
    
    % 确保文件扩展名为.sldd
    if isempty(fileExt)
        slddFullPath = [slddFullPath, '.sldd'];
    elseif ~strcmpi(fileExt, '.sldd')
        slddFullPath = [slddFullPath(1:end-length(fileExt)), '.sldd'];
    end
    
    [dictPath, dictName, dictExt] = fileparts(slddFullPath);
    
    fprintf('\n%s\n', repmat('=', 1, 60));
    fprintf('开始保存到数据字典\n');
    fprintf('%s\n', repmat('=', 1, 60));
    fprintf('数据字典名称: %s%s\n', dictName, dictExt);
    fprintf('保存位置: %s\n', dictPath);
    fprintf('完整路径: %s\n', slddFullPath);
    fprintf('%s\n', repmat('-', 1, 60));
    
    % 获取基础工作区所有变量名
    try
        allVars = evalin('base', 'who');
    catch ME
        error('无法访问基础工作区: %s', ME.message);
    end
    
    if isempty(allVars)
        fprintf('基础工作区没有变量可保存\n');
        return;
    end
    
    fprintf('基础工作区中找到 %d 个变量\n', length(allVars));
    
    % 默认排除列表（系统变量和临时变量）
    excludeVars = {'ans', 'defaultExclude', 'excludeList', 'allVars', 'slddFile', 'filePath', 'fileName', 'fileExt',...
        'addedCount','dataSect','desc','dictEntries','dictObj','errorCount','excludeVars','i','updatedCount',...
        'varClass','varName','varNames','varSize','varValue','verifyDict','verifySect','ME','entryObj',...
        'currentFolder', 'dictPath', 'dictName', 'dictExt', 'slddFullPath'};
    
    % 过滤变量
    varNames = setdiff(allVars, excludeVars);
    
    if isempty(varNames)
        fprintf('所有变量都在排除列表中，没有变量可保存\n');
        return;
    end
    
    fprintf('将保存 %d 个变量（排除了 %d 个变量）\n', length(varNames), length(allVars) - length(varNames));
    
    % 显示将要保存的变量
    fprintf('\n将要保存的变量列表:\n');
    fprintf('%s\n', repmat('-', 1, 40));
    for i = 1:min(length(varNames), 20)  % 只显示前20个，避免过多输出
        try
            varValue = evalin('base', varNames{i});
            varSize = size(varValue);
            varClass = class(varValue);
            
            if isa(varValue, 'Simulink.Parameter') || isa(varValue, 'myPackage.Parameter')
                % 参数对象特殊处理
                if ~isempty(varValue.Description)
                    desc = varValue.Description;
                else
                    desc = '参数对象';
                end
                fprintf('  %2d. %-25s [%s] - %s\n', i, varNames{i}, varClass, desc);
            elseif isa(varValue, 'Simulink.Bus')
                fprintf('  %2d. %-25s [%s] - Bus对象\n', i, varNames{i}, varClass);
            elseif isa(varValue, 'Simulink.Signal')
                fprintf('  %2d. %-25s [%s] - Signal对象\n', i, varNames{i}, varClass);
            else
                fprintf('  %2d. %-25s [%s] %s\n', i, varNames{i}, varClass, mat2str(varSize));
            end
        catch
            fprintf('  %2d. %-25s [无法获取信息]\n', i, varNames{i});
        end
    end
    
    if length(varNames) > 20
        fprintf('  ... 还有 %d 个变量未显示\n', length(varNames) - 20);
    end
    fprintf('%s\n', repmat('-', 1, 40));
    
    % 询问用户是否继续
    choice = input('是否继续保存到数据字典？(Y/N, 默认Y): ', 's');
    if ~isempty(choice) && upper(choice) == 'N'
        fprintf('用户取消保存操作\n');
        return;
    end
    
    % 打开或创建数据字典
    try
        if exist(slddFullPath, 'file')
            dictObj = Simulink.data.dictionary.open(slddFullPath);
            fprintf('打开现有数据字典: %s\n', dictName);
        else
            dictObj = Simulink.data.dictionary.create(slddFullPath);
            fprintf('创建新数据字典: %s\n', dictName);
        end
    catch ME
        error('无法打开/创建数据字典: %s', ME.message);
    end
   
    % 获取数据字典的设计数据段
    try
        dataSect = getSection(dictObj, 'Design Data');
    catch ME
        close(dictObj);
        error('无法获取数据字典设计数据段: %s', ME.message);
    end
     
    % 统计信息
    addedCount = 0;
    updatedCount = 0;
    errorCount = 0;
    unchangedCount = 0;
    
    % 遍历并保存每个变量
    fprintf('\n正在保存变量...\n');
    for i = 1:length(varNames)
        varName = varNames{i};
        
        try
            varValue = evalin('base', varName);
            if hasEntry(dataSect, varName)
                % 更新现有条目
                entryObj = getEntry(dataSect, varName);
                currentValue = getValue(entryObj);
                
                % 检查值是否发生变化
                if isequal(varValue, currentValue)
                    unchangedCount = unchangedCount + 1;
                    if i <= 10  % 只显示前10个未变化的变量
                        fprintf('  %s: 无变化，跳过\n', varName);
                    end
                    continue;
                end
                
                setValue(entryObj, varValue);
                updatedCount = updatedCount + 1;
                fprintf('  ✓ %s: 已更新\n', varName);
            else
                % 创建新条目
                addEntry(dataSect, varName, varValue);
                addedCount = addedCount + 1;
                fprintf('  ✓ %s: 已添加\n', varName);
            end

        catch ME
            errorCount = errorCount + 1;
            fprintf('  ✗ %s: 保存失败 - %s\n', varName, ME.message);
        end
    end
    
    % 保存更改到数据字典文件
    try
        saveChanges(dictObj);
        fprintf('\n数据字典更改已保存\n');
    catch ME
        fprintf('无法保存数据字典更改: %s\n', ME.message);
    end
    
    % 关闭数据字典
    try
        close(dictObj);
    catch ME
        fprintf('关闭数据字典时出错: %s\n', ME.message);
    end
    
    % 显示总结信息
    fprintf('\n%s\n', repmat('=', 1, 60));
    fprintf('保存完成！\n');
    fprintf('%s\n', repmat('=', 1, 60));
    fprintf('数据字典信息：\n');
    fprintf('  名称: %s%s\n', dictName, dictExt);
    fprintf('  路径: %s\n', dictPath);
    fprintf('  完整路径: %s\n', slddFullPath);
    fprintf('\n保存统计：\n');
    fprintf('  总变量数: %d\n', length(varNames));
    fprintf('  成功添加: %d\n', addedCount);
    fprintf('  成功更新: %d\n', updatedCount);
    fprintf('  无变化跳过: %d\n', unchangedCount);
    fprintf('  失败数量: %d\n', errorCount);
    
    if errorCount == 0
        fprintf('✓ 所有变量保存成功\n');
    else
        fprintf('⚠ 有 %d 个变量保存失败\n', errorCount);
    end
    
    % 验证保存结果
    fprintf('\n验证保存结果...\n');
    try
        verifyDict = Simulink.data.dictionary.open(slddFullPath);
        verifySect = getSection(verifyDict, 'Design Data');
        dictEntries = find(verifySect);
        fprintf('数据字典中现有条目数: %d\n', length(dictEntries));
        
        % 显示部分条目
        if length(dictEntries) > 0
            fprintf('示例条目：\n');
            for i = 1:min(5, length(dictEntries))
                entryName = dictEntries(i).Name;
                entryValue = dictEntries(i).getValue;
                if isa(entryValue, 'Simulink.Parameter')
                    fprintf('  %s: Simulink.Parameter\n', entryName);
                elseif isa(entryValue, 'Simulink.Bus')
                    fprintf('  %s: Simulink.Bus\n', entryName);
                elseif isa(entryValue, 'Simulink.Signal')
                    fprintf('  %s: Simulink.Signal\n', entryName);
                else
                    fprintf('  %s: %s\n', entryName, class(entryValue));
                end
            end
            if length(dictEntries) > 5
                fprintf('  ... 还有 %d 个条目\n', length(dictEntries) - 5);
            end
        end
        
        close(verifyDict);
    catch ME
        fprintf('验证数据字典时出错: %s\n', ME.message);
    end
    
    fprintf('\n数据字典已成功保存！\n');
    fprintf('您可以在以下位置找到它：\n');
    fprintf('%s\n', slddFullPath);
    
    % 辅助函数: 检查条目是否存在
    function exists = hasEntry(section, entryName)
        try
            getEntry(section, entryName);  % 尝试获取条目
            exists = true;                 % 如果成功，说明条目存在
        catch
            exists = false;                % 如果失败，说明条目不存在
        end
    end
end