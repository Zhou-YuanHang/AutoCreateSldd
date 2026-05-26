function m_to_sldd(slddFile)
%% m_to_sldd — 将基础工作区中的 Simulink 对象保存到数据字典
%  自动执行 xxx_objects.m，将 Signal/Parameter/Bus 等对象写入 .sldd。
%  已有条目会对比值，无变化则跳过。
%
%  用法：
%    m_to_sldd                     % 自动查找 / 执行并保存
%    m_to_sldd('MyDict.sldd')      % 指定数据字典文件名

%% 参数处理
if nargin == 0
    currDir = pwd;
    mFiles = dir(fullfile(currDir, '*_objects.m'));
    if isempty(mFiles)
        error('当前目录未找到 *_objects.m 文件，请先运行 excel_to_sldd 生成。');
    end
    mFile = mFiles(1).name;
    [~, baseName, ~] = fileparts(mFile);
    if endsWith(baseName, '_objects')
        baseName = baseName(1:end-8);
    end
    slddFile = fullfile(currDir, [baseName, '.sldd']);
    fprintf('找到脚本: %s\n', mFile);
    fprintf('数据字典: %s\n\n', slddFile);

    choice = input(sprintf('是否执行 %s 以创建工作区对象？(Y/N, 默认 Y): ', mFile), 's');
    if isempty(choice) || upper(choice) == 'Y'
        run(mFile);
        fprintf('脚本已执行，对象已创建到工作区。\n\n');
    end
else
    [p, n, e] = fileparts(slddFile);
    if isempty(e)
        slddFile = fullfile(p, [n, '.sldd']);
    end
end

%% 获取工作区变量
allVars = evalin('base', 'who');
if isempty(allVars)
    fprintf('基础工作区没有变量。\n');
    return;
end

% 只筛选 Simulink 数据对象
simulinkTypes = {'Simulink.Signal', 'Simulink.Parameter', 'Simulink.Bus'};
varNames = {};
for i = 1:length(allVars)
    try
        v = evalin('base', allVars{i});
        if any(strcmp(class(v), simulinkTypes))
            varNames{end+1} = allVars{i};
        end
    catch
        % 跳过无法读取的变量
    end
end

if isempty(varNames)
    fprintf('工作区中未找到 Simulink.Signal / Parameter / Bus 对象。\n');
    return;
end

fprintf('\n找到 %d 个 Simulink 对象:\n', length(varNames));
for i = 1:length(varNames)
    v = evalin('base', varNames{i});
    fprintf('  %s [%s]\n', varNames{i}, class(v));
end

choice = input('\n是否保存到数据字典？(Y/N, 默认 Y): ', 's');
if ~isempty(choice) && upper(choice) == 'N'
    fprintf('已取消。\n');
    return;
end

%% 打开或创建数据字典
fprintf('\n正在处理...\n');
newSldd = ~exist(slddFile, 'file');
if newSldd
    dictObj = Simulink.data.dictionary.create(slddFile);
    fprintf('创建新数据字典: %s\n', slddFile);
else
    dictObj = Simulink.data.dictionary.open(slddFile);
    fprintf('打开现有数据字典: %s\n', slddFile);
end

dataSect = getSection(dictObj, 'Design Data');

% 只读取本字典自身的条目（不遍历外部引用）
ownEntries = find(dataSect, 'IncludeReferences', false);
ownNames = {ownEntries.Name};

added   = 0;
updated = 0;
skipped = 0;
failed  = 0;

for i = 1:length(varNames)
    vn = varNames{i};
    try
        newVal = evalin('base', vn);
        idx = find(strcmp(ownNames, vn));

        if isempty(idx)
            % 新增条目
            addEntry(dataSect, vn, newVal);
            added = added + 1;
            fprintf('  + %s (新增)\n', vn);
        else
            % 已有条目，对比是否变化
            entryObj = ownEntries(idx);
            oldVal = getValue(entryObj);
            if isequal(newVal, oldVal)
                skipped = skipped + 1;
            else
                setValue(entryObj, newVal);
                updated = updated + 1;
                fprintf('  ~ %s (已修改)\n', vn);
            end
        end
    catch ME
        failed = failed + 1;
        fprintf('  x %s: %s\n', vn, ME.message);
    end
end

saveChanges(dictObj);
close(dictObj);

%% 输出统计
fprintf('\n===== 保存完成 =====\n');
fprintf('数据字典: %s\n', slddFile);
fprintf('  新增:  %d\n', added);
fprintf('  修改:  %d\n', updated);
fprintf('  跳过(无变化): %d\n', skipped);
if failed > 0
    fprintf('  失败:  %d\n', failed);
end
end
