function m_to_sldd(slddFile)
%% m_to_sldd — 将基础工作区中的 Simulink 对象保存到数据字典
%  用法：
%    m_to_sldd                     % 自动查找并执行 *_objects.m，再写入 .sldd
%    m_to_sldd('MyDict.sldd')      % 指定数据字典文件名

%% ========== 参数处理 ==========
if nargin == 0
    currDir = pwd;
    mFiles = dir(fullfile(currDir, '*_objects.m'));
    if isempty(mFiles)
        error('当前目录未找到 *_objects.m 文件，请先运行 excel_to_sldd 生成。');
    end
    mFilePath = fullfile(currDir, mFiles(1).name);
    [~, baseName, ~] = fileparts(mFilePath);
    if endsWith(baseName, '_objects')
        baseName = baseName(1:end-8);
    end
    slddFile = fullfile(currDir, [baseName, '.sldd']);
    [~, mName] = fileparts(mFilePath);
    fprintf('找到脚本: %s\n', mName);
    fprintf('数据字典: %s\n\n', slddFile);

    choice = input(sprintf('是否执行 %s 以创建工作区对象？(Y/N, 默认 Y): ', mName), 's');
    if isempty(choice) || upper(choice) == 'Y'
        % 在基础工作区执行脚本，变量不会随函数退出销毁
        evalin('base', sprintf('run(''%s'')', mFilePath));
        fprintf('脚本已执行，对象已创建到工作区。\n\n');
    end
end

%% ========== 列出所有存在的 .sldd 文件（诊断用） ==========
fprintf('[调试] 搜索当前目录已有 .sldd 文件...\n');
existingSldd = dir('*.sldd');
if isempty(existingSldd)
    fprintf('[调试]   未找到已有 .sldd 文件\n');
else
    for i = 1:length(existingSldd)
        fprintf('[调试]   找到: %s\n', existingSldd(i).name);
    end
end
fprintf('[调试] 目标 SLDD 路径: %s\n', slddFile);
fprintf('[调试] 文件是否存在: %d\n', exist(slddFile, 'file'));

%% ========== 获取基础工作区变量 ==========
fprintf('\n[调试] 读取基础工作区变量...\n');
allVars = evalin('base', 'who');
fprintf('[调试]   基础工作区共有 %d 个变量\n', length(allVars));

% 筛选 Simulink 数据对象（用 isa 判断继承，支持 myPackage.Signal 等子类）
simulinkTypes = {'Simulink.Signal', 'Simulink.Parameter', 'Simulink.Bus'};
varNames = {};
for i = 1:length(allVars)
    try
        v = evalin('base', allVars{i});
        c = class(v);
        fprintf('[调试]   变量 %s: 类 = %s\n', allVars{i}, c);
        isSimulinkObj = false;
        for j = 1:length(simulinkTypes)
            if isa(v, simulinkTypes{j})
                isSimulinkObj = true;
                break;
            end
        end
        if isSimulinkObj
            varNames{end+1} = allVars{i};
            fprintf('[调试]     -> 已选中\n');
        end
    catch ME
        fprintf('[调试]   变量 %s: 读取失败 - %s\n', allVars{i}, ME.message);
    end
end

if isempty(varNames)
    fprintf('工作区中未找到 Simulink.Signal / Parameter / Bus 对象。\n');
    return;
end
fprintf('\n共找到 %d 个 Simulink 对象\n', length(varNames));

choice = input('\n是否保存到数据字典？(Y/N, 默认 Y): ', 's');
if ~isempty(choice) && upper(choice) == 'N'
    fprintf('已取消。\n');
    return;
end

%% ========== 打开或创建数据字典 ==========
fprintf('\n[调试] 打开/创建数据字典...\n');
try
    if exist(slddFile, 'file')
        dictObj = Simulink.data.dictionary.open(slddFile);
        fprintf('打开现有数据字典: %s\n', slddFile);
    else
        dictObj = Simulink.data.dictionary.create(slddFile);
        fprintf('创建新数据字典: %s\n', slddFile);
    end
catch ME
    error('数据字典操作失败: %s', ME.message);
end

fprintf('[调试] 获取 Design Data 段...\n');
try
    dataSect = getSection(dictObj, 'Design Data');
catch ME
    error('获取 Design Data 段失败: %s', ME.message);
end

fprintf('[调试] 查找本字典已有条目（用 getEntry 试错，比 find 更可靠）...\n');

%% ========== 逐条处理 ==========
added   = 0;
updated = 0;
skipped = 0;
failed  = 0;

fprintf('\n正在写入条目...\n');
for i = 1:length(varNames)
    vn = varNames{i};
    newVal = evalin('base', vn);

    % 尝试获取已有条目
    try
        entryObj = getEntry(dataSect, vn);
        % 条目已存在 → 对比值
        entryExists = true;
    catch
        % 条目不存在 → 新增
        entryExists = false;
    end

    try
        if entryExists
            oldVal = getValue(entryObj);
            if isEqualByStruct(newVal, oldVal)
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

%% ========== 保存并关闭 ==========
fprintf('\n[调试] 保存更改...\n');
try
    saveChanges(dictObj);
    fprintf('[调试] saveChanges 成功\n');
catch ME
    fprintf('[调试] saveChanges 失败: %s\n', ME.message);
end

fprintf('[调试] 关闭数据字典...\n');
try
    close(dictObj);
    fprintf('[调试] close 成功\n');
catch ME
    fprintf('[调试] close 失败: %s\n', ME.message);
end

%% ========== 输出统计 ==========
fprintf('\n===== 保存完成 =====\n');
fprintf('数据字典: %s\n', slddFile);
fprintf('  新增:  %d\n', added);
fprintf('  修改:  %d\n', updated);
fprintf('  跳过(无变化): %d\n', skipped);
if failed > 0
    fprintf('  失败:  %d\n', failed);
end

% 验证文件是否生成
if exist(slddFile, 'file')
    fprintf('\n✅ 文件已确认生成: %s\n', slddFile);
else
    fprintf('\n❌ 文件未生成！\n');
end
end


%% ================================================================
function eq = isEqualByStruct(a, b)
if ~strcmp(class(a), class(b))
    eq = false;
    return;
end
try
    sa = struct(a);
    sb = struct(b);
    eq = isequal(sa, sb);
catch
    try
        eq = isequal(a, b);
    catch
        eq = false;
    end
end
end
