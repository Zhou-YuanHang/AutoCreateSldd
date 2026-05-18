function sldd_to_excel()
%% sldd_to_excel — SLDD → Excel 反向同步
%  MATLAB 负责读 SLDD、对比数据，Python 负责写 Excel（保留全部格式）。
%
%  用法：在 MATLAB 命令窗口运行 sldd_to_excel

clc;
fprintf('===============================================\n');
fprintf('   数据字典 → Excel 反向同步工具\n');
fprintf('===============================================\n\n');

%% 选择文件
[filename, filepath] = uigetfile({'*.sldd', '数据字典 (*.sldd)'}, '选择数据字典');
if isequal(filename, 0), disp('已取消'); return; end
slddPath = fullfile(filepath, filename);
[~, baseName, ~] = fileparts(filename);
excelPath = fullfile(filepath, [baseName, '.xlsx']);
fprintf('SLDD: %s\n', slddPath);
fprintf('Excel: %s\n', excelPath);

%% 读取 SLDD
dictObj = Simulink.data.dictionary.open(slddPath);
entries = find(getSection(dictObj, 'Design Data'));
close(dictObj);
fprintf('找到 %d 个条目\n', length(entries));

%% 分类
data = struct('Signal',{}, 'Parameter',{}, 'Const',{}, 'Bus',{}, 'BusElement',{});
data(1).Signal = {}; data.Parameter = {}; data.Const = {}; data.Bus = {}; data.BusElement = {};

for i = 1:length(entries)
    en = entries(i);
    v = getValue(en);
    if isa(v, 'Simulink.Bus')
        data.Bus{end+1} = en;
        try; for e = 1:length(v.Elements)
                data.BusElement{end+1} = {en.Name, v.Elements(e)};
            end; catch; end
    elseif isa(v, 'Simulink.Parameter')
        try; sc = v.CoderInfo.CustomStorageClass; catch; sc = ''; end
        if strcmpi(sc, 'Define')
            data.Const{end+1} = en;
        else
            data.Parameter{end+1} = en;
        end
    elseif isa(v, 'Simulink.Signal')
        data.Signal{end+1} = en;
    end
end

%%% 构建 JSON 数据
fprintf('\n--- 写入 Excel ---\n');
jsonData = struct();
jsonData.excelPath = excelPath;
jsonData.sheets = struct();
stats.added = 0;

sheetList = {'Signal','Parameter','Const','Bus','BusElement'};
for si = 1:length(sheetList)
    sn = sheetList{si};
    entries = data.(sn);
    if isempty(entries), continue; end

    headers = sheetHeaders(sn);
    reqCols = requiredCols(sn);
    newRows = buildRows(entries, sn);

    % 全量输出：Python 侧负责同名更新、无名追加
    stats.added = stats.added + size(newRows, 1);

    % 全部列（含选填）用 SLDD 值覆盖
    outRows = cell(size(newRows,1), 1);
    for i = 1:size(newRows, 1)
        row = newRows(i, :);
        % 转换数值类型为字符串
        for c = 1:length(row)
            if isnumeric(row{c})
                if isempty(row{c}); row{c} = '';
                elseif isnan(row{c}); row{c} = '';
                elseif isscalar(row{c}); row{c} = num2str(row{c});
                else; row{c} = mat2str(row{c}); end
            elseif islogical(row{c})
                if row{c}; row{c} = 'True'; else; row{c} = 'False'; end
            end
        end
        outRows{i} = row;
    end

    sheetObj = struct();
    sheetObj.headers = {headers{:}};
    sheetObj.requiredCols = reqCols;
    sheetObj.rows = {outRows{:}};
    jsonData.sheets.(sn) = sheetObj;
end

%% 读取目录下的 Enum .m 文件，加入 JSON
enumDir = filepath;
enumFiles = dir(fullfile(enumDir, '*.m'));
enumRows = {};
for fi = 1:length(enumFiles)
    fn = enumFiles(fi).name;
    % 跳过 _objects.m 和本文件
    if contains(fn, '_objects.m') || strcmp(fn, 'sldd_to_excel.m') ...
            || strcmp(fn, 'sldd_to_excel_helper.m')
        continue;
    end
    ep = fullfile(enumDir, fn);
    try
        txt = fileread(ep);
        % 提取 classdef Name < BaseType
        tk = regexp(txt, 'classdef\s+(\w+)\s*<\s*(\w+)', 'tokens');
        if isempty(tk), continue; end
        className = tk{1}{1};
        baseType  = tk{1}{2};

        % 提取枚举成员 Label (Value)
        mems = regexp(txt, '^\s+(\w+)\s*\(([^)]+)\)\s*$', 'tokens', 'lineanchors');
        for mi = 1:length(mems)
            enumRows{end+1, 1} = {className, mems{mi}{1}, str2double(mems{mi}{2}), baseType};
        end
        fprintf('  Enum: %s (%d 成员)\n', className, length(mems));
    catch
    end
end

if ~isempty(enumRows)
    fprintf('  Enum: %d 条\n', length(enumRows));
    headers = {'EnumName','EnumNumbers','Value','DataType'};
    % 用和 Signal 一致的结构：{row1, row2, ...}
    outRows = cell(length(enumRows), 1);
    for i = 1:length(enumRows)
        row = enumRows{i};
        % 转字符串
        for c = 1:length(row)
            if isnumeric(row{c}); row{c} = num2str(row{c}); end
        end
        outRows{i} = row;
    end
    sheetObj = struct();
    sheetObj.headers = {headers{:}};
    sheetObj.requiredCols = [1 2 3 4];
    sheetObj.rows = {outRows{:}};
    jsonData.sheets.Enum = sheetObj;
    stats.added = stats.added + length(enumRows);
end

%% 写入 JSON，调 Python
if isempty(fieldnames(jsonData.sheets))
    fprintf('  无变化，无需写入\n');
    return;
end

jsonPath = fullfile(tempdir, 'sldd_to_excel_temp.json');
fid = fopen(jsonPath, 'w', 'n', 'UTF-8');
fprintf(fid, '%s', jsonencode(jsonData));
fclose(fid);

% 找 Python 路径
pyExe = '';
% 1. 尝试 pyenv（MATLAB 配置的 Python）
try
    pe = pyenv;
    if pe.Executable ~= ""
        [ok, ~] = system(sprintf('"%s" --version >nul 2>nul', pe.Executable));
        if ok == 0; pyExe = pe.Executable; end
    end
catch, end

% 2. 尝试系统 PATH
if isempty(pyExe)
    [ok, ~] = system('python --version >nul 2>nul');
    if ok == 0; pyExe = 'python'; end
end
if isempty(pyExe)
    [ok, ~] = system('python3 --version >nul 2>nul');
    if ok == 0; pyExe = 'python3'; end
end

% 3. 尝试常见安装路径
commonPaths = {
    'D:\Python\python.exe';
    'C:\Python39\python.exe';
    'C:\Python310\python.exe';
    'C:\Python311\python.exe';
    'C:\Python312\python.exe';
    'C:\Program Files\Python39\python.exe';
    'C:\Program Files\Python310\python.exe';
    'C:\Program Files\Python311\python.exe';
    'C:\Program Files\Python312\python.exe';
    };
for i = 1:length(commonPaths)
    if exist(commonPaths{i}, 'file')
        pyExe = commonPaths{i};
        break;
    end
end

if isempty(pyExe)
    error('找不到 Python。请将 Python 安装路径添加到系统 PATH 环境变量后重试。');
end
fprintf('使用 Python: %s\n', pyExe);

scriptPath = fullfile(fileparts(mfilename('fullpath')), 'sldd_to_excel_helper.py');
cmd = sprintf('"%s" "%s" "%s"', pyExe, scriptPath, jsonPath);
fprintf('\n--- 写入 Excel ---\n');
[ret, out] = system(cmd);
fprintf('%s', out);

if ret ~= 0
    error('Python 写入失败 (exit=%d)', ret);
end

delete(jsonPath);

fprintf('\n===== 同步完成 =====\n');
fprintf('Excel: %s\n', excelPath);
fprintf('  已同步: %d 条\n', stats.added);
end


%% ================================================================
function h = sheetHeaders(sn)
switch sn
    case 'Signal'
        h = {'VariableName','Package','Object','CustomStorageClass','DataType',...
             'InitialValue','HeaderFile','DefinitionFile',...
             'Description','Min','Max','Unit','Dimensions','Complexity'};
    case 'Parameter'
        h = {'VariableName','Package','Object','CustomStorageClass','DataType',...
             'InitialValue','HeaderFile','DefinitionFile',...
             'Description','Min','Max','Unit','Dimensions','Complexity'};
    case 'Const'
        h = {'Name','Value','DataType','HeaderFile'};
    case 'Bus'
        h = {'BusName','Description','HeaderFile','Alignment','PreserveElementDimensions','DataScope'};
    case 'BusElement'
        h = {'BusName','ElementName','DataType','Dimensions','Description','Unit'};
end
end

function cols = requiredCols(sn)
switch sn
    case 'Signal';       cols = [1 2 3 4 5 7 8];
    case 'Parameter';    cols = [1 2 3 4 5 6 7 8];
    case 'Const';        cols = [1 2 3 4];
    case 'Bus';          cols = [1 2 3 4 5 6];
    case 'BusElement';   cols = [1 2 3 4];
end
end

function rows = buildRows(entries, sn)
h = sheetHeaders(sn);
rows = cell(length(entries), length(h));
for i = 1:length(entries)
    switch sn
        case 'Signal';     rows(i,:) = getSignalRow(entries{i});
        case 'Parameter';  rows(i,:) = getParamRow(entries{i});
        case 'Const';      rows(i,:) = getConstRow(entries{i});
        case 'Bus';        rows(i,:) = getBusRow(entries{i});
        case 'BusElement'; rows(i,:) = getBusElemRow(entries{i});
    end
end
end

function row = getSignalRow(en)
v = getValue(en);
try; hf = v.CoderInfo.CustomAttributes.HeaderFile; catch; hf = ''; end
try; df = v.CoderInfo.CustomAttributes.DefinitionFile; catch; df = ''; end
row = {en.Name, extractPkg(v), 'Signal', readSC(v), strVal(v.DataType), ...
       strVal(v.InitialValue), hf, df, strVal(v.Description), ...
       numVal(v.Min), numVal(v.Max), strVal(v.DocUnits), numVal(v.Dimensions), strVal(v.Complexity)};
end

function row = getParamRow(en)
v = getValue(en);
try; hf = v.CoderInfo.CustomAttributes.HeaderFile; catch; hf = ''; end
try; df = v.CoderInfo.CustomAttributes.DefinitionFile; catch; df = ''; end
row = {en.Name, extractPkg(v), 'Parameter', readSC(v), strVal(v.DataType), ...
       numVal(v.Value), hf, df, strVal(v.Description), ...
       numVal(v.Min), numVal(v.Max), strVal(v.DocUnits), numVal(v.Dimensions), strVal(v.Complexity)};
end

function row = getConstRow(en)
v = getValue(en);
try; hf = v.CoderInfo.CustomAttributes.HeaderFile; catch; hf = ''; end
row = {en.Name, numVal(v.Value), strVal(v.DataType), hf};
end

function row = getBusRow(en)
v = getValue(en);
row = {en.Name, strVal(v.Description), strVal(v.HeaderFile), ...
       numVal(v.Alignment), boolVal(v.PreserveElementDimensions), strVal(v.DataScope)};
end

function row = getBusElemRow(cd)
e = cd{2};
row = {cd{1}, e.Name, strVal(e.DataType), numVal(e.Dimensions), ...
       strVal(e.Description), strVal(e.DocUnits)};
end

function s = toCmp(v)
if isempty(v); s = ''; return; end
if isnumeric(v)
    if isnan(v); s = ''; elseif numel(v)==1; s = num2str(v,12); else; s = mat2str(v); end
elseif islogical(v); if v; s='true'; else; s='false'; end
elseif isstring(v); s = char(v);
elseif ischar(v); s = strtrim(lower(v));
else; s = strtrim(lower(char(v)));
end
end

function pkg = extractPkg(v)
c = class(v); d = strfind(c, '.');
if isempty(d); pkg = 'Simulink'; else; pkg = c(1:d(1)-1); end
end

function sc = readSC(v)
def = {'Auto','SimulinkGlobal','ExportedGlobal','ImportedExtern','ImportedExternPointer'};
try
    sc = v.CoderInfo.StorageClass;
    if any(strcmp(sc, def)); return; end
    if strcmp(sc, 'Custom')
        try; sc = v.CoderInfo.CustomStorageClass; catch; sc = 'Custom'; end
        return;
    end
    if strcmp(sc, 'Default'); sc = ''; end
catch; sc = ''; end
end

function s = strVal(v)
if isempty(v); s = ''; return; end
if islogical(v); if v; s='true'; else; s='false'; end; return; end
s = char(v);
end

function s = numVal(v)
if isempty(v); s = ''; return; end
if isnumeric(v)
    if numel(v)==1; s = num2str(v); else; s = mat2str(v); end
elseif islogical(v); if v; s='true'; else; s='false'; end
else; s = char(v);
end
end

function s = boolVal(v)
if isempty(v); s = ''; return; end
if v; s = 'True'; else; s = 'False'; end
end
