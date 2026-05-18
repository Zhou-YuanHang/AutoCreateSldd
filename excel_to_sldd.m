function excel_to_sldd()
%% excel_to_sldd — 一键完成：读 Excel → 生成 M 脚本 → 执行 → 写入 .sldd
%  按 F5 运行，不污染基础工作区。
%  自动找当前目录 .xlsx → 选 sheet → 生成 xxx_objects.m
%  → 在基础工作区执行 → 对比并写入 .sldd（保留外部引用）

clc;
fprintf('===============================================\n');
fprintf('   Excel 数据字典导入工具（全自动）\n');
fprintf('===============================================\n\n');

%% ==================== 第一步：选择 Excel 文件 ====================
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

%% ==================== 第二步：获取 Sheet 列表 ====================
[~, AllSheets] = xlsfinfo(fullpath);
fprintf('找到的工作表: %s\n', strjoin(AllSheets, ', '));

[~, baseName, ~] = fileparts(filename);

%% ==================== 第三步：生成 xxx_objects.m ====================
outputFilename = fullfile(filepath, [baseName, '_objects.m']);
fprintf('\n--- 生成 M 脚本 ---\n');
Excel2Workspace(AllSheets, fullpath, outputFilename, filename);

if ~exist(outputFilename, 'file')
    error('M 脚本生成失败！');
end
fprintf('M 脚本已生成: %s\n', outputFilename);

%% ==================== 第四步：执行 M 脚本 ====================
fprintf('\n--- 执行 M 脚本 ---\n');
choice = input(sprintf('是否执行 %s_objects.m 以创建工作区对象？(Y/N, 默认 Y): ', baseName), 's');
if isempty(choice) || upper(choice) == 'Y'
    evalin('base', sprintf('run(''%s'')', outputFilename));
    fprintf('脚本已执行，对象已创建到工作区。\n');
end

%% ==================== 第五步：写入 .sldd ====================
slddFile = fullfile(filepath, [baseName, '.sldd']);
fprintf('\n--- 写入数据字典 ---\n');

allVars = evalin('base', 'who');
if isempty(allVars)
    error('基础工作区没有变量，请先执行 M 脚本。');
end

% 筛选 Simulink 对象（支持 myPackage.Signal 等子类）
simulinkTypes = {'Simulink.Signal', 'Simulink.Parameter', 'Simulink.Bus'};
varNames = {};
for i = 1:length(allVars)
    try
        v = evalin('base', allVars{i});
        isSimObj = false;
        for j = 1:length(simulinkTypes)
            if isa(v, simulinkTypes{j})
                isSimObj = true;
                break;
            end
        end
        if isSimObj
            varNames{end+1} = allVars{i};
        end
    catch
    end
end

if isempty(varNames)
    error('工作区中未找到 Simulink.Signal / Parameter / Bus 对象。');
end
fprintf('找到 %d 个 Simulink 对象\n', length(varNames));

% 创建或打开 SLDD
if exist(slddFile, 'file')
    dictObj = Simulink.data.dictionary.open(slddFile);
    fprintf('打开现有数据字典: %s\n', slddFile);
else
    dictObj = Simulink.data.dictionary.create(slddFile);
    fprintf('创建新数据字典: %s\n', slddFile);
end
dataSect = getSection(dictObj, 'Design Data');

% 逐条处理
added   = 0;
updated = 0;
skipped = 0;
failed  = 0;

for i = 1:length(varNames)
    vn = varNames{i};
    newVal = evalin('base', vn);

    try
        entryObj = getEntry(dataSect, vn);
        entryExists = true;
    catch
        entryExists = false;
    end

    try
        if entryExists
            oldVal = getValue(entryObj);
            if isSameValue(newVal, oldVal)
                skipped = skipped + 1;
                fprintf('  - %s (无变化，跳过)\n', vn);
            else
                setValue(entryObj, newVal);
                updated = updated + 1;
                fprintf('  ~ %s (已修改)\n', vn);
            end
        else
            addEntry(dataSect, vn, newVal);
            added = added + 1;
            fprintf('  + %s (新增)\n', vn);
        end
    catch ME
        failed = failed + 1;
        fprintf('  x %s: %s\n', vn, ME.message);
    end
end

saveChanges(dictObj);
close(dictObj);

fprintf('\n===== 完成 =====\n');
fprintf('数据字典: %s\n', slddFile);
fprintf('  新增:  %d\n', added);
fprintf('  修改:  %d\n', updated);
fprintf('  跳过(无变化): %d\n', skipped);
if failed > 0
    fprintf('  失败:  %d\n', failed);
end
fprintf('\n✅ 所有步骤已完成！\n');
end


%% ================================================================
function same = isSameValue(a, b)
% 转 struct 后比较，绕开 handle 类比较陷阱
if ~strcmp(class(a), class(b))
    same = false;
    return;
end
try
    same = isequal(struct(a), struct(b));
catch
    try
        same = isequal(a, b);
    catch
        same = false;
    end
end
end
