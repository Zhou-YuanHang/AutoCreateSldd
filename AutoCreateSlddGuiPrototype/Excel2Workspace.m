function Excel2Workspace(AllSheets, fullpath, outputFilename, filename, selection)
% EXCEL2WORKSPACE 将Excel数据字典转换为MATLAB对象初始化脚本
%
% 输入参数：
%   AllSheets      - Excel中所有工作表名称列表（cell数组）
%   fullpath       - Excel文件完整路径
%   outputFilename - 输出 .m 脚本的完整路径
%   filename       - Excel文件名（用于 detectImportOptions）
%   selection      - 可选，显式指定要处理的工作表索引或名称列表
%
% 支持的工作表类型：Signal / Parameter / Bus / BusElement

% ------------------------------------------------------------------
% 必要列定义
% ------------------------------------------------------------------
REQUIRED_COLUMNS = struct( ...
    'Signal',     {{'VariableName','Package','Object','CustomStorageClass','DataType','HeaderFile','DefinitionFile'}}, ...
    'Parameter',  {{'VariableName','Package','Object','CustomStorageClass','DataType','HeaderFile','DefinitionFile','InitialValue'}}, ...
    'Bus',        {{'BusName','Description','HeaderFile','Alignment','PreserveElementDimensions','DataScope'}}, ...
    'BusElement', {{'BusName','ElementName','DataType','Dimensions'}} ...
    );

% ------------------------------------------------------------------
% 工作表选择
% ------------------------------------------------------------------
% 兼容原 CLI 交互模式，也允许 GUI 显式传入选择结果，避免二次弹框或选择丢失
if nargin < 5 || isempty(selection)
    % 按名称预选已知的标准工作表，避免硬编码序号
    knownSheets   = {'Signal', 'Parameter', 'Bus', 'BusElement'};
    defaultSelect = find(ismember(AllSheets, knownSheets));
    if isempty(defaultSelect)
        defaultSelect = 1;
    end

    [selection, ok] = listdlg( ...
        'PromptString',  '选择要读取的工作表(可多选)', ...
        'SelectionMode', 'multiple', ...
        'ListString',    AllSheets, ...
        'Name',          '工作表选择', ...
        'ListSize',      [300, 150], ...
        'InitialValue',  defaultSelect);

    if ~ok
        fprintf('用户取消了选择工作表。\n');
        return;
    end
elseif iscell(selection) || isstring(selection)
    selectedNames = cellstr(string(selection));
    selection = zeros(1, numel(selectedNames));
    for idx = 1:numel(selectedNames)
        matchIdx = find(strcmp(AllSheets, selectedNames{idx}), 1, 'first');
        if isempty(matchIdx)
            error('指定的工作表不存在：%s', selectedNames{idx});
        end
        selection(idx) = matchIdx;
    end
elseif ~isnumeric(selection)
    error('selection 必须是工作表索引或名称列表。');
end

selection = unique(selection(:).', 'stable');
if isempty(selection)
    error('没有可处理的工作表。');
end
if any(selection < 1 | selection > numel(AllSheets) | mod(selection, 1) ~= 0)
    error('selection 中存在非法的工作表索引。');
end

% ------------------------------------------------------------------
% 强制保证 Bus 在 BusElement 之前处理（避免依赖顺序导致元素被静默跳过）
% ------------------------------------------------------------------
selectedNames = AllSheets(selection);
busIdx        = find(strcmp(selectedNames, 'Bus'), 1, 'first');
busElemIdx    = find(strcmp(selectedNames, 'BusElement'), 1, 'first');
if ~isempty(busIdx) && ~isempty(busElemIdx) && busIdx > busElemIdx
    selection([busElemIdx, busIdx]) = selection([busIdx, busElemIdx]);
    fprintf('注意：已自动调整处理顺序，Bus 将在 BusElement 之前处理。\n');
end


% ------------------------------------------------------------------
% 创建输出文件
% ------------------------------------------------------------------
fid = fopen(outputFilename, 'w');
if fid == -1
    error('无法创建输出文件: %s', outputFilename);
end
fprintf(fid, '%% 自动生成的数据对象文件\n');
fprintf(fid, '%% 创建时间: %s\n', datetime('now'));
fprintf(fid, '%% 来源Excel: %s\n\n', fullpath);

% 用于存储已处理的 Bus 名称（供 BusElement 验证使用）
busInfo = containers.Map();

% ------------------------------------------------------------------
% 逐工作表处理
% ------------------------------------------------------------------
for sheetIdx = selection
    ProcessingSheet = AllSheets{sheetIdx};
    fprintf('\n正在处理的sheet: %s\n', ProcessingSheet);

    switch ProcessingSheet

        case 'Bus'
            processBusSheet(fid, fullpath, filename, ProcessingSheet, ...
                REQUIRED_COLUMNS.Bus, busInfo);

        case 'BusElement'
            processBusElementSheet(fid, fullpath, filename, ProcessingSheet, ...
                REQUIRED_COLUMNS.BusElement, busInfo);

        case 'Signal'
            processCoderSheet(fid, fullpath, filename, ProcessingSheet, ...
                REQUIRED_COLUMNS.Signal, 'Signal');

        case 'Parameter'
            processCoderSheet(fid, fullpath, filename, ProcessingSheet, ...
                REQUIRED_COLUMNS.Parameter, 'Parameter');

        otherwise
            fprintf('跳过未知工作表: %s\n', ProcessingSheet);
    end
end

% 关闭输出文件
fclose(fid);
fprintf('\n=============== 导入完成 ===============\n');
fprintf('数据对象已保存到文件: %s\n', outputFilename);
fprintf('要使用这些对象，请在MATLAB中运行此脚本文件。\n');
end


% ==================================================================
%  子函数：处理 Bus 工作表
% ==================================================================
function processBusSheet(fid, fullpath, filename, sheetName, requiredCols, busInfo)
    try
        opts = detectImportOptions(filename, 'Sheet', sheetName);
        opts = setvartype(opts, {'BusName','Description','HeaderFile','DataScope'}, 'string');
        opts = setvartype(opts, {'Alignment'}, 'double');
        opts = setvartype(opts, {'PreserveElementDimensions'}, 'logical');
        dataTable = readtable(fullpath, opts);
    catch ME
        fprintf('读取工作表 %s 出错: %s\n', sheetName, ME.message);
        return;
    end

    fprintf('成功读取工作表 %s，共 %d 行 %d 列\n', sheetName, size(dataTable,1), size(dataTable,2));
    columnNames = dataTable.Properties.VariableNames;
    fprintf('工作表列名: %s\n', strjoin(columnNames, ', '));

    [ok, missing] = checkRequiredColumns(requiredCols, columnNames);
    if ~ok
        fprintf('工作表 %s 缺少必要列: %s，跳过整张表\n', sheetName, strjoin(missing, ', '));
        return;
    end
    fprintf('所有必须列都存在，开始处理数据\n');
    fprintf(fid, '\n%%%% ================= Bus 对象 ===================%%\n\n');

    totalRows = height(dataTable);
    fprintf('Bus对象数量：%d 个\n', totalRows);

    for row = 1:totalRows
        currentRow = dataTable(row, :);
        busNameCol = getActualColumnName('BusName', columnNames);
        busName    = string(currentRow.(busNameCol));

        if isempty(busName) || ismissing(busName)
            fprintf('跳过行 %d：Bus名称为空\n', row + 1);
            continue;
        end

        printProgress(row, totalRows, string(busName));

        descriptionCol  = getActualColumnName('Description',              columnNames);
        headerFileCol   = getActualColumnName('HeaderFile',               columnNames);
        alignmentCol    = getActualColumnName('Alignment',                columnNames);
        preserveDimsCol = getActualColumnName('PreserveElementDimensions',columnNames);
        dataScopeCol    = getActualColumnName('DataScope',                columnNames);

        description  = string(currentRow.(descriptionCol));
        headerFile   = string(currentRow.(headerFileCol));
        alignment    = currentRow.(alignmentCol);
        preserveDims = currentRow.(preserveDimsCol);
        dataScope    = string(currentRow.(dataScopeCol));

        try
            fprintf(fid, '%% ----- Bus对象: %s -----\n', busName);
            fprintf(fid, '%s = Simulink.Bus;\n', busName);

            if hasNonEmptyValue(description)
                fprintf(fid, "%s.Description = '%s';\n", busName, description);
            else
                fprintf(fid, "%s.Description = '';\n", busName);
            end
            if hasNonEmptyValue(headerFile)
                fprintf(fid, "%s.HeaderFile = '%s';\n", busName, headerFile);
            else
                fprintf(fid, "%s.HeaderFile = '';\n", busName);
            end
            fprintf(fid, '%s.Alignment = %d;\n', busName, alignment);
            fprintf(fid, '%s.PreserveElementDimensions = %s;\n', busName, string(logical(preserveDims)));
            if hasNonEmptyValue(dataScope)
                fprintf(fid, "%s.DataScope = '%s';\n", busName, dataScope);
            else
                fprintf(fid, "%s.DataScope = 'Auto';\n", busName);
            end
            fprintf(fid, '\n');

            busInfo(busName) = struct();
        catch ME
            warning('处理Bus %s 失败: %s (行 %d)', busName, ME.message, row + 1);
        end
    end
    fprintf('完成对Bus对象的处理\n');
end


% ==================================================================
%  子函数：处理 BusElement 工作表
% ==================================================================
function processBusElementSheet(fid, fullpath, filename, sheetName, requiredCols, busInfo)
    try
        opts = detectImportOptions(filename, 'Sheet', sheetName);
        opts = setvartype(opts, 'string');
        dataTable = readtable(fullpath, opts);
    catch ME
        fprintf('读取工作表 %s 出错: %s\n', sheetName, ME.message);
        return;
    end

    fprintf('成功读取工作表 %s，共 %d 行 %d 列\n', sheetName, size(dataTable,1), size(dataTable,2));
    columnNames = dataTable.Properties.VariableNames;
    fprintf('工作表列名: %s\n', strjoin(columnNames, ', '));

    [ok, missing] = checkRequiredColumns(requiredCols, columnNames);
    if ~ok
        fprintf('工作表 %s 缺少必要列: %s，跳过整张表\n', sheetName, strjoin(missing, ', '));
        return;
    end
    fprintf('所有必须列都存在，开始处理数据\n');
    fprintf(fid, '\n%%%% ================= BusElement 定义 ===================%%\n\n');

    busNameCol     = getActualColumnName('BusName',     columnNames);
    elementNameCol = getActualColumnName('ElementName', columnNames);
    dataTypeCol    = getActualColumnName('DataType',    columnNames);
    dimensionsCol  = getActualColumnName('Dimensions',  columnNames);

    % 可选列
    dimensionsModeCol = getOptionalColumn('DimensionsMode', columnNames);
    complexityCol     = getOptionalColumn('Complexity',     columnNames);
    samplingModeCol   = getOptionalColumn('SamplingMode',   columnNames);
    minCol            = getOptionalColumn('Min',            columnNames);
    maxCol            = getOptionalColumn('Max',            columnNames);
    docUnitsCol       = getOptionalColumn('DocUnits',       columnNames);
    if isempty(docUnitsCol)
        docUnitsCol   = getOptionalColumn('Unit',           columnNames);
    end
    descriptionCol    = getOptionalColumn('Description',    columnNames);

    busNames  = unique(dataTable.(busNameCol));
    totalRows = height(dataTable);
    fprintf('BusElement 元素数量：%d 个\n', totalRows);

    for b = 1:length(busNames)
        busName = string(busNames(b));
        if ismissing(busName) || strlength(busName) == 0
            continue;
        end

        if ~isKey(busInfo, busName)
            fprintf('警告: Bus "%s" 未在Bus工作表中定义，跳过其元素（请确保同时选中了Bus工作表）\n', busName);
            continue;
        end

        busElements  = dataTable(strcmp(dataTable.(busNameCol), busName), :);
        numElements  = height(busElements);
        if numElements == 0; continue; end

        fprintf(fid, '%% ----- 为Bus "%s" 创建元素 -----\n', busName);
        tempVarName = sprintf('%s_elements', matlab.lang.makeValidName(busName));
        fprintf(fid, '%s = Simulink.BusElement;\n', tempVarName);

        for elemIdx = 1:numElements
            currentElement = busElements(elemIdx, :);
            elementName    = string(currentElement.(elementNameCol));
            fprintf(fid, '%% 元素 %d: %s\n', elemIdx, elementName);

            fprintf(fid, '%s(%d).Name = ''%s'';\n', tempVarName, elemIdx, elementName);

            % Description
            fprintf(fid, '%s(%d).Description = ''%s'';\n', tempVarName, elemIdx, ...
                getStringValue(currentElement, descriptionCol, ''));

            % DataType
            fprintf(fid, '%s(%d).DataType = ''%s'';\n', tempVarName, elemIdx, ...
                getStringValue(currentElement, dataTypeCol, 'double'));

            % Dimensions
            writeDimensions(fid, tempVarName, elemIdx, currentElement, dimensionsCol);

            % DimensionsMode
            fprintf(fid, '%s(%d).DimensionsMode = ''%s'';\n', tempVarName, elemIdx, ...
                getStringValue(currentElement, dimensionsModeCol, 'Fixed'));

            % Complexity
            fprintf(fid, '%s(%d).Complexity = ''%s'';\n', tempVarName, elemIdx, ...
                getStringValue(currentElement, complexityCol, 'real'));

            % SamplingMode
            fprintf(fid, '%s(%d).SamplingMode = ''%s'';\n', tempVarName, elemIdx, ...
                getStringValue(currentElement, samplingModeCol, 'Sample based'));

            % Min / Max
            fprintf(fid, '%s(%d).Min = %s;\n', tempVarName, elemIdx, ...
                getNumericValue(currentElement, minCol, '[]'));
            fprintf(fid, '%s(%d).Max = %s;\n', tempVarName, elemIdx, ...
                getNumericValue(currentElement, maxCol, '[]'));

            % DocUnits
            fprintf(fid, '%s(%d).DocUnits = ''%s'';\n', tempVarName, elemIdx, ...
                getStringValue(currentElement, docUnitsCol, ''));

            fprintf(fid, '\n');
        end

        fprintf(fid, '%% 将元素数组赋值给Bus对象\n');
        fprintf(fid, '%s.Elements = %s;\n', busName, tempVarName);
        fprintf(fid, 'clear %s;\n\n', tempVarName);
        fprintf('为Bus ''%s'' 添加了 %d 个元素\n', busName, numElements);
    end
    fprintf('完成对BusElement的处理\n');
end


% ==================================================================
%  子函数：统一处理 Signal / Parameter 工作表
%  sheetType: 'Signal' | 'Parameter'
% ==================================================================
function processCoderSheet(fid, fullpath, filename, sheetName, requiredCols, sheetType)
    try
        opts = detectImportOptions(filename, 'Sheet', sheetName);
        opts = setvartype(opts, {'VariableName','Package','Object','CustomStorageClass', ...
            'DataType','HeaderFile','DefinitionFile','InitialValue','Unit','Description'}, 'string');
        opts = setvartype(opts, {'Min','Max'}, 'double');
        dataTable = readtable(fullpath, opts);
    catch ME
        fprintf('读取工作表 %s 出错: %s\n', sheetName, ME.message);
        return;
    end

    fprintf('成功读取工作表 %s，共 %d 行 %d 列\n', sheetName, size(dataTable,1), size(dataTable,2));
    columnNames = dataTable.Properties.VariableNames;
    fprintf('工作表列名: %s\n', strjoin(columnNames, ', '));

    % 列检查统一采用大小写不敏感方式（与 Bus/BusElement 一致）
    [ok, missing] = checkRequiredColumns(requiredCols, columnNames);
    if ~ok
        fprintf('工作表 %s 缺少必要列: %s，跳过整张表\n', sheetName, strjoin(missing, ', '));
        return;
    end
    fprintf('所有必须列都存在，开始处理数据\n');

    switch sheetType
        case 'Signal'
            sectionTitle = 'Signal 对象';
            objLabel     = '信号对象';
        case 'Parameter'
            sectionTitle = 'Parameter 对象';
            objLabel     = '参数对象';
    end

    fprintf(fid, '\n%%%% ===== %s =====\n\n', sectionTitle);
    totalRows = height(dataTable);
    fprintf('%s 数量：%d 个\n', objLabel, totalRows);

    varNameCol        = getActualColumnName('VariableName',       columnNames);
    packageCol        = getActualColumnName('Package',            columnNames);
    objectTypeCol     = getActualColumnName('Object',             columnNames);
    storageClassCol   = getActualColumnName('CustomStorageClass', columnNames);
    dataTypeCol       = getActualColumnName('DataType',           columnNames);
    headerFileCol     = getActualColumnName('HeaderFile',         columnNames);
    definitionFileCol = getActualColumnName('DefinitionFile',     columnNames);
    initialValueCol   = getOptionalColumn('InitialValue',         columnNames);

    defaultStorageClasses = {'Auto','SimulinkGlobal','ExportedGlobal','ImportedExtern','ImportedExternPointer'};

    for row = 1:totalRows
        currentRow  = dataTable(row, :);
        varName     = string(currentRow.(varNameCol));

        if ismissing(varName) || strlength(varName) == 0
            fprintf('跳过行 %d：变量名为空\n', row + 1);
            continue;
        end

        printProgress(row, totalRows, varName);

        % 读取必要列
        package        = string(currentRow.(packageCol));
        objectType     = string(currentRow.(objectTypeCol));
        storageClass   = string(currentRow.(storageClassCol));
        dataType       = string(currentRow.(dataTypeCol));
        headerFile     = string(currentRow.(headerFileCol));
        definitionFile = string(currentRow.(definitionFileCol));

        % InitialValue（仅 Parameter 有效）
        if strcmp(sheetType, 'Parameter') && ~isempty(initialValueCol)
            initialValue = string(currentRow.(initialValueCol));
        else
            initialValue = missing; % Signal 不写初始值
        end

        % 可选数值列
        minVal = getTableNumeric(currentRow, columnNames, 'Min', '[]');
        maxVal = getTableNumeric(currentRow, columnNames, 'Max', '[]');

        % 可选字符串列
        unit        = getTableString(currentRow, columnNames, 'Unit',        "");
        description = getTableString(currentRow, columnNames, 'Description', "");

        try
            requestedCtor = sprintf('%s.%s', package, objectType);
            fallbackCtor  = sprintf('Simulink.%s', sheetType);

            fprintf(fid, '%% ----- %s: %s -----\n', objLabel, varName);
            writeSafeConstructor(fid, varName, requestedCtor, fallbackCtor);

            % 存储类
            writeStorageClassAssignment(fid, varName, storageClass, defaultStorageClasses);

            % 头文件 / 定义文件
            writeCoderAttributeAssignment(fid, varName, 'HeaderFile', headerFile, false);
            writeCoderAttributeAssignment(fid, varName, 'DefinitionFile', definitionFile, true);

            % 数据类型
            if hasNonEmptyValue(dataType)
                fprintf(fid, '%s.DataType = ''%s'';\n', varName, escapeMatlabString(dataType));
            end

            % 初始值（仅 Parameter）
            if strcmp(sheetType, 'Parameter') && hasNonEmptyValue(initialValue)
                fprintf(fid, '%s.Value = %s;\n', varName, initialValue);
            end

            % 最小/最大值
            fprintf(fid, '%s.Min = %s;\n', varName, minVal);
            fprintf(fid, '%s.Max = %s;\n', varName, maxVal);

            % 描述
            if hasNonEmptyValue(description)
                fprintf(fid, '%s.Description = ''%s'';\n', varName, escapeMatlabString(description));
            else
                fprintf(fid, '%s.Description = '''';\n', varName);
            end

            % 单位
            if hasNonEmptyValue(unit)
                fprintf(fid, '%s.DocUnits = ''%s'';\n', varName, escapeMatlabString(unit));
            else
                fprintf(fid, '%s.DocUnits = '''';\n', varName);
            end

            fprintf(fid, '\n');
        catch ME
            warning('处理变量 %s 失败: %s (行 %d)', varName, ME.message, row);
        end

    end
    fprintf('完成对 %s 变量的处理\n', sheetType);
end


% ==================================================================
%  工具函数
% ==================================================================

% 检查列名是否存在（大小写不敏感，忽略空格）
function [allExist, missingList] = checkRequiredColumns(requiredCols, columnNames)
    missingList = {};
    for k = 1:length(requiredCols)
        if ~checkColumnExists(requiredCols{k}, columnNames)
            missingList{end+1} = requiredCols{k}; %#ok<AGROW>
        end
    end
    allExist = isempty(missingList);
end

function exists = checkColumnExists(colName, columnNames)
    normName  = strrep(colName,     ' ', '');
    normNames = strrep(columnNames, ' ', '');
    exists    = any(strcmpi(normName, normNames));
end

function actualName = getActualColumnName(colName, columnNames)
    if any(strcmpi(colName, columnNames))
        actualName = colName;
        return;
    end
    normName  = strrep(colName,     ' ', '');
    normNames = strrep(columnNames, ' ', '');
    for i = 1:length(columnNames)
        if strcmpi(normName, normNames{i})
            actualName = columnNames{i};
            return;
        end
    end
    actualName = colName; % 找不到时返回原名，后续报错提示
end

% 返回可选列的实际列名（不存在时返回空字符串）
function colName = getOptionalColumn(name, columnNames)
    if checkColumnExists(name, columnNames)
        colName = getActualColumnName(name, columnNames);
    else
        colName = '';
    end
end

% 判断值是否为非空、非缺失内容，并始终返回标量逻辑值
function tf = hasNonEmptyValue(v)
    if iscell(v)
        if isempty(v)
            tf = false;
            return;
        end
        v = v{1};
    end

    if isempty(v)
        tf = false;
        return;
    end

    if isnumeric(v) || islogical(v)
        tf = any(~ismissing(v(:)));
        return;
    end

    s = string(v);
    tf = any(~ismissing(s(:)) & strlength(strtrim(s(:))) > 0);
end

% 从 table 行中读取字符串值，列不存在或为缺失时返回默认值
function val = getStringValue(row, colName, defaultVal)
    if isempty(colName)
        val = defaultVal;
        return;
    end
    v = row.(colName);
    if hasNonEmptyValue(v)
        val = char(string(v));
    else
        val = defaultVal;
    end
end

% 从 table 行中读取数值字符串，列不存在或为缺失时返回默认值
function val = getNumericValue(row, colName, defaultVal)
    if isempty(colName)
        val = defaultVal;
        return;
    end
    v = row.(colName);
    if hasNonEmptyValue(v)
        val = char(string(v));
    else
        val = defaultVal;
    end
end

% 从已知 columnNames 中获取字符串列（用于 Signal/Parameter）
function val = getTableString(row, columnNames, colName, defaultVal)
    actualCol = getOptionalColumn(colName, columnNames);
    if ~isempty(actualCol) && hasNonEmptyValue(row.(actualCol))
        val = string(row.(actualCol));
    else
        val = defaultVal;
    end
end

% 从已知 columnNames 中获取数值列并返回字符串表达
function val = getTableNumeric(row, columnNames, colName, defaultVal)
    actualCol = getOptionalColumn(colName, columnNames);
    if ~isempty(actualCol) && hasNonEmptyValue(row.(actualCol))
        val = string(row.(actualCol));
    else
        val = defaultVal;
    end
end

function writeSafeConstructor(fid, varName, requestedCtor, fallbackCtor)
    requestedCtor = char(string(requestedCtor));
    fallbackCtor = char(string(fallbackCtor));

    if strcmp(requestedCtor, fallbackCtor)
        fprintf(fid, '%s = %s;\n', varName, requestedCtor);
        return;
    end

    fprintf(fid, 'try\n');
    fprintf(fid, '    %s = %s;\n', varName, requestedCtor);
    fprintf(fid, 'catch ME\n');
    fprintf(fid, '    warning(''AutoCreateSldd:ConstructorFallback'', ''变量 %s 的构造器 %s 不可用，已回退为 %s。原因: %%s'', ME.message);\n', ...
        escapeMatlabString(varName), escapeMatlabString(requestedCtor), escapeMatlabString(fallbackCtor));
    fprintf(fid, '    %s = %s;\n', varName, fallbackCtor);
    fprintf(fid, 'end\n');
end

function writeStorageClassAssignment(fid, varName, storageClass, defaultStorageClasses)
    if ~hasNonEmptyValue(storageClass)
        fprintf(fid, '%s.CoderInfo.StorageClass = ''Auto'';\n', varName);
        return;
    end

    storageClass = char(string(storageClass));
    if any(strcmp(storageClass, defaultStorageClasses))
        fprintf(fid, '%s.CoderInfo.StorageClass = ''%s'';\n', varName, escapeMatlabString(storageClass));
        return;
    end

    fprintf(fid, 'try\n');
    fprintf(fid, '    %s.CoderInfo.StorageClass = ''Custom'';\n', varName);
    fprintf(fid, '    %s.CoderInfo.CustomStorageClass = ''%s'';\n', varName, escapeMatlabString(storageClass));
    fprintf(fid, 'catch ME\n');
    fprintf(fid, '    warning(''AutoCreateSldd:StorageClassFallback'', ''变量 %s 的自定义存储类 %s 在当前环境不可用，已回退为 Auto。原因: %%s'', ME.message);\n', ...
        escapeMatlabString(varName), escapeMatlabString(storageClass));
    fprintf(fid, '    try\n');
    fprintf(fid, '        %s.CoderInfo.StorageClass = ''Auto'';\n', varName);
    fprintf(fid, '    catch\n');
    fprintf(fid, '    end\n');
    fprintf(fid, 'end\n');
end

function writeCoderAttributeAssignment(fid, varName, attrName, attrValue, required)
    if nargin < 5
        required = false;
    end
    if ~required && ~hasNonEmptyValue(attrValue)
        return;
    end

    if hasNonEmptyValue(attrValue)
        attrLiteral = escapeMatlabString(attrValue);
    else
        attrLiteral = '';
    end

    fprintf(fid, 'if isprop(%s.CoderInfo.CustomAttributes, ''%s'')\n', varName, attrName);
    fprintf(fid, '    try\n');
    fprintf(fid, '        %s.CoderInfo.CustomAttributes.%s = ''%s'';\n', varName, attrName, attrLiteral);
    fprintf(fid, '    catch ME\n');
    fprintf(fid, '        warning(''AutoCreateSldd:CoderAttributeSkipped'', ''变量 %s 的 %s 写入失败，已跳过。原因: %%s'', ME.message);\n', ...
        escapeMatlabString(varName), escapeMatlabString(attrName));
    fprintf(fid, '    end\n');
    fprintf(fid, 'else\n');
    fprintf(fid, '    warning(''AutoCreateSldd:CoderAttributeSkipped'', ''变量 %s 当前存储类不支持 %s，已跳过。'');\n', ...
        escapeMatlabString(varName), escapeMatlabString(attrName));
    fprintf(fid, 'end\n');
end

function text = escapeMatlabString(value)
    text = char(string(value));
    text = strrep(text, '''', '''''');
end

% 写入 Dimensions 字段，支持数组格式 [x,y] 和标量
function writeDimensions(fid, tempVarName, elemIdx, currentElement, dimensionsCol)

    if isempty(dimensionsCol)
        fprintf(fid, '%s(%d).Dimensions = 1;\n', tempVarName, elemIdx);
        return;
    end
    v = currentElement.(dimensionsCol);
    if iscell(v); v = v{1}; end
    dimsStr = string(v);
    if ismissing(dimsStr) || strlength(dimsStr) == 0
        fprintf(fid, '%s(%d).Dimensions = 1;\n', tempVarName, elemIdx);
        return;
    end
    try
        if startsWith(dimsStr, '[') && endsWith(dimsStr, ']')
            fprintf(fid, '%s(%d).Dimensions = %s;\n', tempVarName, elemIdx, dimsStr);
        else
            dimNum = str2double(dimsStr);
            if ~isnan(dimNum)
                fprintf(fid, '%s(%d).Dimensions = %d;\n', tempVarName, elemIdx, dimNum);
            else
                fprintf(fid, '%s(%d).Dimensions = %s;\n', tempVarName, elemIdx, dimsStr);
            end
        end
    catch
        fprintf(fid, '%s(%d).Dimensions = 1;\n', tempVarName, elemIdx);
    end
end

% 进度打印（按10%间隔或最后一行）
function printProgress(row, totalRows, label)
    interval = max(1, floor(totalRows / 10));
    if mod(row, interval) == 0 || row == totalRows
        fprintf('处理中: %d/%d (%.0f%%) - %s\n', ...
            row, totalRows, 100 * row / totalRows, label);
    end
end
