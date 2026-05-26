function Excel2Workspace(AllSheets, fullpath, outputFilename, filename)
% clear;clc;

% [filename,filepath] = uigetfile({'*.xlsx;*.xls', 'Excel文件 (*.xlsx, *.xls)'},'选择数据字典Excel文件','WS2Excel_Template.xlsx');

% %检查用户是否取消选择
% if isequal(filename,0)
%     disp("用户取消选择excel文件")
% end
%
% % 构建完整文件路径
% fullpath = fullfile(filepath,filename);
%
% %检查文件是否存在
% if ~exist(fullpath,"file")
%     error("错误：文件不存在或无法访问。请检查文件路径和权限")
% else
%     fprintf("文件存在\n");
% end
%
%
% %检查文件是否可读
% [fid,message] = fopen(fullpath,'r');
% if fid  == -1
%     error("错误：无法读取文件。原因：%s",message)
% else
%     fclose(fid);
%     fprintf("文件可读\n");
% end
%
%
% % 获取Excel文件中的工作表信息
% try
%     sheets = sheetnames(fullpath);
%     if isempty(sheets)
%         error('无法读取Excel文件，请确保文件格式正确且未被损坏。');
%     end
%     fprintf('找到的工作表: %s\n', strjoin(sheets, ', '));
% catch ME
%     fprintf('读取Excel文件信息时出错:\n');
%     fprintf('  错误信息: %s\n', ME.message);
%     return;
% end

% 让用户选择要读取的工作表
[selection, ok] = listdlg("PromptString", "选择要读取的工作表(可多选)", ...
    "SelectionMode", 'multiple', ...
    "ListString", AllSheets, ...
    'Name', '工作表选择', ...
    "ListSize", [300, 150], ...
    "InitialValue", [2 3]);
if ~ok
    fprintf("用户取消了选择工作表。\n");
    return;
end

% 检测需要的必要列 - 新增Bus和BusElement
REQUIRED_COLUMNS = struct( ...
    "Signal", {{'VariableName','Package','Object','CustomStorageClass','DataType','HeaderFile'}}, ...
    "Parameter", {{'VariableName','Package','Object','CustomStorageClass','DataType','HeaderFile','InitialValue'}}, ...
    "Const", {{'VariableName','Package','Object','CustomStorageClass','DataType','HeaderFile','InitialValue'}}, ...
    "Bus", {{'BusName','Description','HeaderFile','Alignment','PreserveElementDimensions','DataScope'}}, ...
    "BusElement", {{'BusName','ElementName','DataType','Dimensions'}} ...
    );

% 辅助函数：检查列名是否存在（不区分大小写和空格）
    function columnExists = checkColumnExists(columnName, columnNames)
        % 去除空格和转换为小写后进行比较
        normalizedColumnName = lower(strrep(columnName, ' ', ''));
        normalizedColumnNames = lower(strrep(columnNames, ' ', ''));
        columnExists = any(strcmp(normalizedColumnName, normalizedColumnNames));
    end

% 辅助函数：获取实际的列名（处理大小写和空格问题）
    function actualColumnName = getActualColumnName(columnName, columnNames)
        % 如果直接存在，直接返回
        if any(strcmp(columnName, columnNames))
            actualColumnName = columnName;
            return;
        end

        % 尝试不区分大小写和空格的匹配
        normalizedColumnName = lower(strrep(columnName, ' ', ''));
        normalizedColumnNames = lower(strrep(columnNames, ' ', ''));

        for i = 1:length(columnNames)
            if strcmp(normalizedColumnName, normalizedColumnNames{i})
                actualColumnName = columnNames{i};
                return;
            end
        end

        % 如果没找到，返回原始列名（可能会导致错误）
        actualColumnName = columnName;
    end

% %创建输出.m文件
% outputFilename = strrep(filename,'.xlsx','.m');
% if isempty(outputFilename)
%     outputFilename = 'generated_objects.m';
% else
%     outputFilename = [strrep(filename, '.xls', ''), '_objects.m'];
% end

fid = fopen(outputFilename, 'w');
if fid == -1
    error('无法创建输出文件: %s', outputFilename);
end
fprintf(fid, '%% 自动生成的数据对象文件\n');
fprintf(fid, '%% 创建时间: %s\n', datetime("now"));
fprintf(fid, '%% 来源Excel: %s\n\n', fullpath);

% 用于存储Bus信息
busInfo = containers.Map();

% 处理所有选中的工作表
for sheetIdx = selection
    ProcessingSheet = AllSheets{sheetIdx};

    fprintf("\n正在处理的sheet: %s\n", ProcessingSheet);

    % 处理Bus工作表
    if strcmp(ProcessingSheet, 'Bus')
        try
            opts = detectImportOptions(filename, "Sheet", ProcessingSheet);
            opts = setvartype(opts, {'BusName','Description','HeaderFile','DataScope'}, 'string');
            opts = setvartype(opts, {'Alignment'}, 'double');
            opts = setvartype(opts, {'PreserveElementDimensions'}, 'logical');

            dataTable = readtable(fullpath, opts);
            fprintf("成功读取工作表%s，共 %d 行 %d 列\n", ProcessingSheet, size(dataTable,1), size(dataTable,2));

            columnNames = dataTable.Properties.VariableNames;
            fprintf("工作表列名: %s\n", strjoin(columnNames, ', '));

            % 检测必要列是否存在 - 使用灵活检查
            requiredCols = REQUIRED_COLUMNS.(ProcessingSheet);
            missingColumns = {};
            for colIdx = 1:length(requiredCols)
                requiredCol = requiredCols{colIdx};
                if ~checkColumnExists(requiredCol, columnNames)
                    missingColumns{end+1} = requiredCol;
                end
            end

            if ~isempty(missingColumns)
                fprintf('工作表 %s 缺少以下必要列: %s,将跳过%s整张表的处理\n', ProcessingSheet, strjoin(missingColumns, ', '), ProcessingSheet);
                continue;
            else
                fprintf('所有必须列都存在，将开始处理数据\n');

                % 处理Bus数据
                fprintf("\n开始处理Bus数据\n");
                fprintf(fid, '\n%% ================= Bus 对象 ===================%%\n\n');

                totalRows = height(dataTable);
                fprintf("Bus对象数量为：%s个\n", num2str(totalRows));

                for row = 1:totalRows
                    currentRow = dataTable(row, :);

                    % 使用灵活方式获取实际列名
                    busNameCol = getActualColumnName('BusName', columnNames);
                    busName = string(currentRow.(busNameCol));

                    if isempty(busName) || ismissing(busName)
                        fprintf("跳过行 %d: Bus名称为空\n", row+1);
                        continue;
                    end

                    % 显示当前处理状态
                    if mod(row, 10) == 0 || row == totalRows
                        fprintf("处理中: %d/%d (%.1f%%) - %s\n", ...
                            row, totalRows, 100*row/totalRows, string(busName));
                    end

                    % 获取其他列值
                    descriptionCol = getActualColumnName('Description', columnNames);
                    headerFileCol = getActualColumnName('HeaderFile', columnNames);
                    alignmentCol = getActualColumnName('Alignment', columnNames);
                    preserveDimsCol = getActualColumnName('PreserveElementDimensions', columnNames);
                    dataScopeCol = getActualColumnName('DataScope', columnNames);

                    description = string(currentRow.(descriptionCol));
                    headerFile = string(currentRow.(headerFileCol));
                    alignment = currentRow.(alignmentCol);
                    preserveDims = currentRow.(preserveDimsCol);
                    dataScope = string(currentRow.(dataScopeCol));

                    % 检查可选列
                    if checkColumnExists('DocUnits', columnNames)
                        docUnitsCol = getActualColumnName('DocUnits', columnNames);
                        if ~ismissing(currentRow.(docUnitsCol))
                            docUnits = string(currentRow.(docUnitsCol));
                        else
                            docUnits = '';
                        end
                    else
                        docUnits = '';
                    end

                    try
                        % 创建Bus对象
                        fprintf(fid, '%% ----- Bus对象: %s -----\n', busName);
                        fprintf(fid, '%s = Simulink.Bus;\n', busName);

                        % 设置Bus属性
                        if ~ismissing(description)
                            fprintf(fid, "%s.Description = '%s';\n", busName, description);
                        else
                            fprintf(fid, "%s.Description = '';\n", busName);
                        end

                        if ~ismissing(headerFile)
                            fprintf(fid, "%s.HeaderFile = '%s';\n", busName, headerFile);
                        else
                            fprintf(fid, "%s.HeaderFile = '';\n", busName);
                        end

                        fprintf(fid, '%s.Alignment = %d;\n', busName, alignment);
                        fprintf(fid, '%s.PreserveElementDimensions = %s;\n', busName, ...
                            string(logical(preserveDims)));

                        if ~ismissing(dataScope)
                            fprintf(fid, "%s.DataScope = '%s';\n", busName, dataScope);
                        else
                            fprintf(fid, "%s.DataScope = 'Auto';\n", busName);
                        end

                        % 存储Bus信息
                        busInfo(busName) = struct();

                        fprintf(fid, '\n');

                    catch ME
                        warning("处理Bus %s 失败: %s (行 %d)", busName, ME.message, row+1);
                    end
                end
                fprintf("完成对Bus对象的处理\n");
            end
        catch ME
            fprintf('读取工作表出错: %s\n', ME.message);
        end

        % 处理BusElement工作表
    elseif strcmp(ProcessingSheet, 'BusElement')
        try
            opts = detectImportOptions(filename, "Sheet", ProcessingSheet);
            % 读取所有列作为字符串，稍后转换
            opts = setvartype(opts, 'string');

            dataTable = readtable(fullpath, opts);
            fprintf("成功读取工作表%s，共 %d 行 %d 列\n", ProcessingSheet, size(dataTable,1), size(dataTable,2));

            columnNames = dataTable.Properties.VariableNames;
            fprintf("工作表列名: %s\n", strjoin(columnNames, ', '));

            % 检测必要列是否存在 - 使用灵活检查
            requiredCols = REQUIRED_COLUMNS.(ProcessingSheet);
            missingColumns = {};
            for colIdx = 1:length(requiredCols)
                requiredCol = requiredCols{colIdx};
                if ~checkColumnExists(requiredCol, columnNames)
                    missingColumns{end+1} = requiredCol;
                end
            end

            if ~isempty(missingColumns)
                fprintf('工作表 %s 缺少以下必要列: %s,将跳过%s整张表的处理\n', ProcessingSheet, strjoin(missingColumns, ', '), ProcessingSheet);
                continue;
            else
                fprintf('所有必须列都存在，将开始处理数据\n');

                % 处理BusElement数据
                fprintf("\n开始处理BusElement数据\n");
                fprintf(fid, '\n%% ================= BusElement 定义 ===================%%\n\n');

                totalRows = height(dataTable);
                fprintf("BusElement元素数量为：%s个\n", num2str(totalRows));

                % 获取实际的列名
                busNameCol = getActualColumnName('BusName', columnNames);
                elementNameCol = getActualColumnName('ElementName', columnNames);
                dataTypeCol = getActualColumnName('DataType', columnNames);
                dimensionsCol = getActualColumnName('Dimensions', columnNames);

                % 检查其他可选列
                dimensionsModeCol = '';
                if checkColumnExists('DimensionsMode', columnNames)
                    dimensionsModeCol = getActualColumnName('DimensionsMode', columnNames);
                end

                complexityCol = '';
                if checkColumnExists('Complexity', columnNames)
                    complexityCol = getActualColumnName('Complexity', columnNames);
                end

                samplingModeCol = '';
                if checkColumnExists('SamplingMode', columnNames)
                    samplingModeCol = getActualColumnName('SamplingMode', columnNames);
                end

                minCol = '';
                if checkColumnExists('Min', columnNames)
                    minCol = getActualColumnName('Min', columnNames);
                end

                maxCol = '';
                if checkColumnExists('Max', columnNames)
                    maxCol = getActualColumnName('Max', columnNames);
                end

                docUnitsCol = '';
                if checkColumnExists('DocUnits', columnNames)
                    docUnitsCol = getActualColumnName('DocUnits', columnNames);
                end

                descriptionCol = '';
                if checkColumnExists('Description', columnNames)
                    descriptionCol = getActualColumnName('Description', columnNames);
                end

                % 按BusName分组
                busNames = unique(dataTable.(busNameCol));

                for b = 1:length(busNames)
                    busName = string(busNames(b));
                    if ismissing(busName) || strlength(busName) == 0
                        continue;
                    end

                    % 检查Bus是否已定义
                    if ~isKey(busInfo, busName)
                        fprintf('警告: Bus "%s" 未在Bus工作表中定义，将跳过其元素\n', busName);
                        continue;
                    end

                    % 获取该Bus的所有元素
                    busElements = dataTable(strcmp(dataTable.(busNameCol), busName), :);
                    numElements = height(busElements);

                    if numElements == 0
                        continue;
                    end

                    fprintf(fid, '%% ----- 为Bus "%s" 创建元素 -----\n', busName);

                    % 创建临时元素数组变量
                    tempVarName = sprintf('%s_elements', matlab.lang.makeValidName(busName));
                    fprintf(fid, '%s = Simulink.BusElement;\n', tempVarName);

                    for elemIdx = 1:numElements
                        currentElement = busElements(elemIdx, :);

                        elementName = string(currentElement.(elementNameCol));
                        fprintf(fid, '%% 元素 %d: %s\n', elemIdx, elementName);

                        % 设置元素属性
                        fprintf(fid, '%s(%d).Name = ''%s'';\n', tempVarName, elemIdx, elementName);

                        if ~isempty(descriptionCol) && ~ismissing(currentElement.(descriptionCol))
                            desc = string(currentElement.(descriptionCol));
                            fprintf(fid, '%s(%d).Description = ''%s'';\n', tempVarName, elemIdx, desc);
                        else
                            fprintf(fid, '%s(%d).Description = '''';\n', tempVarName, elemIdx);
                        end

                        if ~ismissing(currentElement.(dataTypeCol))
                            dataType = string(currentElement.(dataTypeCol));
                            fprintf(fid, '%s(%d).DataType = ''%s'';\n', tempVarName, elemIdx, dataType);
                        else
                            fprintf(fid, '%s(%d).DataType = ''double'';\n', tempVarName, elemIdx);
                        end

                        % 处理维度 - 支持多种格式
                        if ~ismissing(currentElement.(dimensionsCol))
                            dims = currentElement.(dimensionsCol);
                            dimsStr = string(dims);

                            % 尝试将维度字符串转换为数值
                            try
                                if startsWith(dimsStr, '[') && endsWith(dimsStr, ']')
                                    % 是数组形式，如 "[3,1]"
                                    fprintf(fid, '%s(%d).Dimensions = %s;\n', tempVarName, elemIdx, dimsStr);
                                else
                                    % 可能是单个数字
                                    dimNum = str2double(dimsStr);
                                    if ~isnan(dimNum)
                                        if dimNum == 1
                                            fprintf(fid, '%s(%d).Dimensions = 1;\n', tempVarName, elemIdx);
                                        else
                                            fprintf(fid, '%s(%d).Dimensions = %d;\n', tempVarName, elemIdx, dimNum);
                                        end
                                    else
                                        fprintf(fid, '%s(%d).Dimensions = %s;\n', tempVarName, elemIdx, dimsStr);
                                    end
                                end
                            catch
                                fprintf(fid, '%s(%d).Dimensions = 1;\n', tempVarName, elemIdx);
                            end
                        else
                            fprintf(fid, '%s(%d).Dimensions = 1;\n', tempVarName, elemIdx);
                        end

                        % 处理维度模式
                        if ~isempty(dimensionsModeCol) && ~ismissing(currentElement.(dimensionsModeCol))
                            dimMode = string(currentElement.(dimensionsModeCol));
                            fprintf(fid, '%s(%d).DimensionsMode = ''%s'';\n', tempVarName, elemIdx, dimMode);
                        else
                            fprintf(fid, '%s(%d).DimensionsMode = ''Fixed'';\n', tempVarName, elemIdx);
                        end

                        % 处理复数类型
                        if ~isempty(complexityCol) && ~ismissing(currentElement.(complexityCol))
                            complexity = string(currentElement.(complexityCol));
                            fprintf(fid, '%s(%d).Complexity = ''%s'';\n', tempVarName, elemIdx, complexity);
                        else
                            fprintf(fid, '%s(%d).Complexity = ''real'';\n', tempVarName, elemIdx);
                        end

                        % 处理采样模式
                        if ~isempty(samplingModeCol) && ~ismissing(currentElement.(samplingModeCol))
                            samplingMode = string(currentElement.(samplingModeCol));
                            fprintf(fid, '%s(%d).SamplingMode = ''%s'';\n', tempVarName, elemIdx, samplingMode);
                        else
                            fprintf(fid, '%s(%d).SamplingMode = ''Sample based'';\n', tempVarName, elemIdx);
                        end

                        % 处理最小/最大值
                        if ~isempty(minCol) && ~ismissing(currentElement.(minCol))
                            minVal = string(currentElement.(minCol));
                            fprintf(fid, '%s(%d).Min = %s;\n', tempVarName, elemIdx, minVal);
                        else
                            fprintf(fid, '%s(%d).Min = [];\n', tempVarName, elemIdx);
                        end

                        if ~isempty(maxCol) && ~ismissing(currentElement.(maxCol))
                            maxVal = string(currentElement.(maxCol));
                            fprintf(fid, '%s(%d).Max = %s;\n', tempVarName, elemIdx, maxVal);
                        else
                            fprintf(fid, '%s(%d).Max = [];\n', tempVarName, elemIdx);
                        end

                        % 处理单位
                        if ~isempty(docUnitsCol) && ~ismissing(currentElement.(docUnitsCol))
                            units = string(currentElement.(docUnitsCol));
                            fprintf(fid, '%s(%d).DocUnits = ''%s'';\n', tempVarName, elemIdx, units);
                        else
                            fprintf(fid, '%s(%d).DocUnits = '''';\n', tempVarName, elemIdx);
                        end

                        fprintf(fid, '\n');
                    end

                    % 将元素数组赋值给Bus对象
                    fprintf(fid, '%% 将元素数组赋值给Bus对象\n');
                    fprintf(fid, '%s.Elements = %s;\n', busName, tempVarName);

                    % 清理临时变量
                    fprintf(fid, 'clear %s;\n\n', tempVarName);

                    fprintf("为Bus '%s' 添加了 %d 个元素\n", busName, numElements);
                end
                fprintf("完成对BusElement的处理\n");
            end
        catch ME
            fprintf('读取工作表出错: %s\n', ME.message);
        end

        % 处理Signal工作表
    elseif strcmp(ProcessingSheet, 'Signal')
        try
            opts = detectImportOptions(filename, "Sheet", ProcessingSheet);
            opts = setvartype(opts, {'InitialValue','VariableName','Unit','Description' ...
                }, 'string');
            opts = setvartype(opts, {'Min','Max'}, 'double');

            dataTable = readtable(fullpath, opts);
            fprintf("成功读取工作表%s，共 %d 行 %d 列\n", ProcessingSheet, size(dataTable,1), size(dataTable,2));

            columnNames = dataTable.Properties.VariableNames;
            fprintf("工作表列名: %s\n", strjoin(columnNames, ', '));

            % 检测必要列是否存在
            requiredCols = REQUIRED_COLUMNS.Signal;
            missingColumns = setdiff(string(requiredCols), string(columnNames));
            if ~isempty(missingColumns)
                fprintf('工作表 %s 缺少以下必要列: %s,将跳过%s整张表的处理\n', ProcessingSheet, strjoin(missingColumns, ', '), ProcessingSheet);
                continue;
            else
                fprintf('所有必须列都存在，将开始处理数据\n');
            end

            fprintf("\n开始处理Signal信号\n");
            fprintf(fid, '\n%% ================= Signal 对象 ===================%%\n\n');

            % 计算需要处理的Signal变量个数
            totalRows = height(dataTable);
            fprintf("Signal信号数量为：%s个\n", num2str(totalRows));

            for row = 1:totalRows
                currentRow = dataTable(row, :);
                currentSignalName = string(currentRow.VariableName);
                if isempty(currentSignalName) || ismissing(currentSignalName)
                    fprintf("跳过行 %d: 变量名为空\n", row+1);
                    continue;
                end

                % 显示当前处理状态
                if mod(row, 10) == 0 || row == totalRows
                    fprintf("处理中: %d/%d (%.1f%%) - %s\n", ...
                        row, totalRows, 100*row/totalRows, string(currentSignalName));
                end

                % 必要列
                package        = string(currentRow.Package);
                objectType     = string(currentRow.Object);
                storageClass   = string(currentRow.CustomStorageClass);
                dataType       = string(currentRow.DataType);
                headerFile     = string(currentRow.HeaderFile);
                definitionFile = string(currentRow.DefinitionFile);

                if ismember('Min', columnNames) && ~ismissing(currentRow.Min)
                    minVal = string(currentRow.Min);
                else
                    minVal = "[]";
                end

                if ismember('Max', columnNames) && ~ismissing(currentRow.Max)
                    maxVal = string(currentRow.Max);
                else
                    maxVal = "[]";
                end

                if ismember('Unit', columnNames) && ~ismissing(currentRow.Unit)
                    unit = string(currentRow.Unit);
                else
                    unit = '';
                end

                if ismember('Description', columnNames) && ~ismissing(currentRow.Description)
                    description = string(currentRow.Description);
                else
                    description = "";
                end

                if ismember('InitialValue', columnNames) && ~ismissing(currentRow.InitialValue)
                    InitialValue = string(currentRow.InitialValue);
                else
                    InitialValue = "";
                end

                try
                    % 创建对象
                    fprintf(fid, '%% ----- 信号对象: %s -----\n', currentSignalName);
                    fprintf(fid, '%s = %s.%s;\n', currentSignalName, package, objectType);

                    % 设置存储类
                    defaultStorageClasses = {'Auto', 'SimulinkGlobal', 'ExportedGlobal', 'ImportedExtern', 'ImportedExternPointer'};
                    if any(strcmp(storageClass, defaultStorageClasses))
                        % 默认存储类
                        fprintf(fid, '%s.CoderInfo.StorageClass = ''%s'';\n', currentSignalName, storageClass);
                    else
                        % 自定义存储类
                        fprintf(fid, '%s.CoderInfo.StorageClass = ''Custom'';\n', currentSignalName);
                        fprintf(fid, '%s.CoderInfo.CustomStorageClass = ''%s'';\n', currentSignalName, storageClass);
                    end

                    % 设置头文件
                    if ~isempty(headerFile) && strlength(headerFile) > 0
                        fprintf(fid, "%s.CoderInfo.CustomAttributes.HeaderFile = '%s';\n", currentSignalName, headerFile);
                    else
                        fprintf(fid, "%s.CoderInfo.CustomAttributes.HeaderFile = '';\n", currentSignalName);
                    end

                    % 设置定义文件
                    if ~isempty(definitionFile) && strlength(definitionFile) > 0
                        fprintf(fid, '%s.CoderInfo.CustomAttributes.DefinitionFile = ''%s'';\n', currentSignalName, definitionFile);
                    else
                        fprintf(fid, '%s.CoderInfo.CustomAttributes.DefinitionFile = '''';\n', currentSignalName);
                    end

                    % 设置数据类型
                    if ~ismissing(dataType)
                        fprintf(fid, '%s.DataType = ''%s'';\n', currentSignalName, dataType);
                    end

                    % Signal没有初始值
                    if ~ismissing(InitialValue) && ~isempty(InitialValue)
                        fprintf(fid, '%s.InitialValue = ''%s'';\n', currentSignalName, InitialValue);
                    end

                    % 设置最小/最大值
                    if ~isempty(minVal) && ~ismissing(minVal)
                        fprintf(fid, '%s.Min = %s;\n', currentSignalName, minVal);
                    end

                    if ~isempty(maxVal) && ~ismissing(maxVal)
                        fprintf(fid, '%s.Max = %s;\n', currentSignalName, maxVal);
                    end

                    % 设置描述
                    if ~ismissing(description)
                        fprintf(fid, '%s.Description = ''%s'';\n', currentSignalName, description);
                    end

                    % 设置单位
                    if ~ismissing(unit)
                        fprintf(fid, "%s.DocUnits = '%s';\n", currentSignalName, unit);
                    else
                        fprintf(fid, "%s.DocUnits = '';\n", currentSignalName);
                    end

                    fprintf(fid, '\n');

                catch ME
                    warning("处理变量 %s 失败: %s (行 %d)", currentSignalName, ME.message, row+1);
                end
            end
            fprintf("完成对Signal变量的处理\n");

        catch ME
            fprintf('读取工作表出错: %s\n', ME.message);
        end

        % 处理Parameter工作表
    elseif strcmp(ProcessingSheet, "Parameter")
        try
            opts = detectImportOptions(filename, "Sheet", ProcessingSheet);
            opts = setvartype(opts, {'InitialValue','VariableName','Unit','Description'}, 'string');
            opts = setvartype(opts, {'Min','Max'}, 'double');

            dataTable = readtable(fullpath, opts);
            fprintf("成功读取工作表%s，共 %d 行 %d 列\n", ProcessingSheet, size(dataTable,1), size(dataTable,2));

            columnNames = dataTable.Properties.VariableNames;
            fprintf("工作表列名: %s\n", strjoin(columnNames, ', '));

            % 检测必要列是否存在
            requiredCols = REQUIRED_COLUMNS.Parameter;
            missingColumns = setdiff(string(requiredCols), string(columnNames));
            if ~isempty(missingColumns)
                fprintf('工作表 %s 缺少以下必要列: %s,将跳过%s整张表的处理\n', ProcessingSheet, strjoin(missingColumns, ', '), ProcessingSheet);
                continue;
            else
                fprintf('所有必须列都存在，将开始处理数据\n');
            end

            fprintf("开始处理Parameter工作表并写入.m文件...\n");
            fprintf(fid, '\n%% ===== Parameter 对象 =====\n\n');
            totalRows = height(dataTable);

            for row = 1:totalRows
                currentRow = dataTable(row, :);
                currentParameterName = string(currentRow.VariableName);

                if ismissing(currentParameterName) || strlength(currentParameterName) == 0
                    fprintf("跳过行 %d: 变量名为空\n", row+1);
                    continue;
                end

                % 显示当前处理状态
                if mod(row, 10) == 0 || row == totalRows
                    fprintf("处理中: %d/%d (%.1f%%) - %s\n", ...
                        row, totalRows, 100*row/totalRows, string(currentParameterName));
                end

                % 必要列
                package        = string(currentRow.Package);
                objectType     = string(currentRow.Object);
                storageClass   = string(currentRow.CustomStorageClass);
                dataType       = string(currentRow.DataType);
                headerFile     = string(currentRow.HeaderFile);
                definitionFile = string(currentRow.DefinitionFile);
                initialValue   = string(currentRow.InitialValue);

                if ismember('Min', columnNames) && ~ismissing(currentRow.Min)
                    minVal = string(currentRow.Min);
                else
                    minVal = "[]";
                end

                if ismember('Max', columnNames) && ~ismissing(currentRow.Max)
                    maxVal = string(currentRow.Max);
                else
                    maxVal = "[]";
                end

                if ismember('Unit', columnNames) && ~ismissing(currentRow.Unit)
                    unit = string(currentRow.Unit);
                else
                    unit = '';
                end

                if ismember('Description', columnNames) && ~ismissing(currentRow.Description)
                    description = string(currentRow.Description);
                else
                    description = "";
                end

                try
                    % 创建对象
                    fprintf(fid, '%% ----- 参数对象: %s -----\n', currentParameterName);
                    fprintf(fid, '%s = %s.%s;\n', currentParameterName, package, objectType);

                    % 设置存储类
                    defaultStorageClasses = {'Auto', 'SimulinkGlobal', 'ExportedGlobal', 'ImportedExtern', 'ImportedExternPointer'};
                    if any(strcmp(storageClass, defaultStorageClasses))
                        % 默认存储类
                        fprintf(fid, '%s.CoderInfo.StorageClass = ''%s'';\n', currentParameterName, storageClass);
                    else
                        % 自定义存储类
                        fprintf(fid, '%s.CoderInfo.StorageClass = ''Custom'';\n', currentParameterName);
                        fprintf(fid, '%s.CoderInfo.CustomStorageClass = ''%s'';\n', currentParameterName, storageClass);
                    end

                    % 设置头文件
                    if ~isempty(headerFile) && strlength(headerFile) > 0
                        fprintf(fid, '%s.CoderInfo.CustomAttributes.HeaderFile = ''%s'';\n', currentParameterName, headerFile);
                    else
                        fprintf(fid, '%s.CoderInfo.CustomAttributes.HeaderFile = '''';\n', currentParameterName);
                    end

                    % 设置定义文件
                    if ~isempty(definitionFile) && strlength(definitionFile) > 0
                        fprintf(fid, '%s.CoderInfo.CustomAttributes.DefinitionFile = ''%s'';\n', currentParameterName, definitionFile);
                    else
                        fprintf(fid, '%s.CoderInfo.CustomAttributes.DefinitionFile = '''';\n', currentParameterName);
                    end

                    % 设置数据类型
                    if ~ismissing(dataType)
                        fprintf(fid, '%s.DataType = ''%s'';\n', currentParameterName, dataType);
                    end

                    % 设置初始值
                    if ~ismissing(initialValue)
                        fprintf(fid, '%s.Value = %s;\n', currentParameterName, initialValue);
                    end

                    % 设置最小/最大值
                    if ~isempty(minVal) && ~ismissing(minVal)
                        fprintf(fid, '%s.Min = %s;\n', currentParameterName, minVal);
                    end

                    if ~isempty(maxVal) && ~ismissing(maxVal)
                        fprintf(fid, '%s.Max = %s;\n', currentParameterName, maxVal);
                    end

                    % 设置描述
                    if ~ismissing(description)
                        fprintf(fid, '%s.Description = ''%s'';\n', currentParameterName, description);
                    end

                    % 设置单位
                    if ~ismissing(unit)
                        fprintf(fid, "%s.DocUnits = '%s';\n", currentParameterName, unit);
                    else
                        fprintf(fid, "%s.DocUnits = '';\n", currentParameterName);
                    end
                    fprintf(fid, '\n');
                catch ME
                    warning("处理变量 %s 失败: %s (行 %d)", currentParameterName, ME.message, row);
                end
            end
            fprintf("完成对Parameter变量的处理\n");

        catch ME
            fprintf('读取工作表出错: %s\n', ME.message);
        end

        % 处理Const工作表
    elseif strcmp(ProcessingSheet, "Const")
        try
            opts = detectImportOptions(filename, "Sheet", ProcessingSheet);
            opts = setvartype(opts, {'InitialValue','VariableName','Unit','Description'}, 'string');
            opts = setvartype(opts, {'Min','Max'}, 'double');

            dataTable = readtable(fullpath, opts);
            fprintf("成功读取工作表%s，共 %d 行 %d 列\n", ProcessingSheet, size(dataTable,1), size(dataTable,2));

            columnNames = dataTable.Properties.VariableNames;
            fprintf("工作表列名: %s\n", strjoin(columnNames, ', '));

            % 检测必要列是否存在
            requiredCols = REQUIRED_COLUMNS.Const;
            missingColumns = setdiff(string(requiredCols), string(columnNames));
            if ~isempty(missingColumns)
                fprintf('工作表 %s 缺少以下必要列: %s,将跳过%s整张表的处理\n', ProcessingSheet, strjoin(missingColumns, ', '), ProcessingSheet);
                continue;
            else
                fprintf('所有必须列都存在，将开始处理数据\n');
            end

            fprintf("开始处理Const工作表并写入.m文件...\n");
            fprintf(fid, '\n%% ===== Const 对象 =====\n\n');
            totalRows = height(dataTable);

            for row = 1:totalRows
                currentRow = dataTable(row, :);
                currentConstName = string(currentRow.VariableName);

                if ismissing(currentConstName) || strlength(currentConstName) == 0
                    fprintf("跳过行 %d: 变量名为空\n", row+1);
                    continue;
                end

                % 显示当前处理状态
                if mod(row, 10) == 0 || row == totalRows
                    fprintf("处理中: %d/%d (%.1f%%) - %s\n", ...
                        row, totalRows, 100*row/totalRows, string(currentConstName));
                end

                % 必要列
                package        = string(currentRow.Package);
                objectType     = string(currentRow.Object);
                storageClass   = string(currentRow.CustomStorageClass);
                dataType       = string(currentRow.DataType);
                headerFile     = string(currentRow.HeaderFile);
                definitionFile = string(currentRow.DefinitionFile);
                initialValue   = string(currentRow.InitialValue);

                if ismember('Min', columnNames) && ~ismissing(currentRow.Min)
                    minVal = string(currentRow.Min);
                else
                    minVal = "[]";
                end

                if ismember('Max', columnNames) && ~ismissing(currentRow.Max)
                    maxVal = string(currentRow.Max);
                else
                    maxVal = "[]";
                end

                if ismember('Unit', columnNames) && ~ismissing(currentRow.Unit)
                    unit = string(currentRow.Unit);
                else
                    unit = '';
                end

                if ismember('Description', columnNames) && ~ismissing(currentRow.Description)
                    description = string(currentRow.Description);
                else
                    description = "";
                end

                try
                    % 创建对象
                    fprintf(fid, '%% ----- 常量对象: %s -----\n', currentConstName);
                    fprintf(fid, '%s = %s.%s;\n', currentConstName, package, objectType);

                    % 设置存储类
                    defaultStorageClasses = {'Auto', 'SimulinkGlobal', 'ExportedGlobal', 'ImportedExtern', 'ImportedExternPointer'};
                    if any(strcmp(storageClass, defaultStorageClasses))
                        % 默认存储类
                        fprintf(fid, '%s.CoderInfo.StorageClass = ''%s'';\n', currentConstName, storageClass);
                    else
                        % 自定义存储类
                        fprintf(fid, '%s.CoderInfo.StorageClass = ''Custom'';\n', currentConstName);
                        fprintf(fid, '%s.CoderInfo.CustomStorageClass = ''%s'';\n', currentConstName, storageClass);
                    end

                    % 设置头文件
                    if ~isempty(headerFile) && strlength(headerFile) > 0
                        fprintf(fid, '%s.CoderInfo.CustomAttributes.HeaderFile = ''%s'';\n', currentConstName, headerFile);
                    else
                        fprintf(fid, '%s.CoderInfo.CustomAttributes.HeaderFile = '''';\n', currentConstName);
                    end

                    % 设置定义文件
                    if ~isempty(definitionFile) && strlength(definitionFile) > 0
                        fprintf(fid, '%s.CoderInfo.CustomAttributes.DefinitionFile = ''%s'';\n', currentConstName, definitionFile);
                    else
                        fprintf(fid, '%s.CoderInfo.CustomAttributes.DefinitionFile = '''';\n', currentConstName);
                    end

                    % 设置数据类型
                    if ~ismissing(dataType)
                        fprintf(fid, '%s.DataType = ''%s'';\n', currentConstName, dataType);
                    end

                    % 设置初始值
                    if ~ismissing(initialValue)
                        fprintf(fid, '%s.Value = %s;\n', currentConstName, initialValue);
                    end

                    % 设置最小/最大值
                    if ~isempty(minVal) && ~ismissing(minVal)
                        fprintf(fid, '%s.Min = %s;\n', currentConstName, minVal);
                    end

                    if ~isempty(maxVal) && ~ismissing(maxVal)
                        fprintf(fid, '%s.Max = %s;\n', currentConstName, maxVal);
                    end

                    % 设置描述
                    if ~ismissing(description)
                        fprintf(fid, '%s.Description = ''%s'';\n', currentConstName, description);
                    end

                    % 设置单位
                    if ~ismissing(unit)
                        fprintf(fid, "%s.DocUnits = '%s';\n", currentConstName, unit);
                    else
                        fprintf(fid, "%s.DocUnits = '';\n", currentConstName);
                    end
                    fprintf(fid, '\n');
                catch ME
                    warning("处理变量 %s 失败: %s (行 %d)", currentConstName, ME.message, row);
                end
            end
            fprintf("完成对Const变量的处理\n");

        catch ME
            fprintf('读取工作表出错: %s\n', ME.message);
        end

    else
        fprintf('跳过未知工作表: %s\n', ProcessingSheet);
    end
end

% 关闭输出文件
fclose(fid);
fprintf('\n=============== 导入完成 ===============\n');
fprintf('数据对象已保存到文件: %s\n', outputFilename);
fprintf('要使用这些对象，请在MATLAB中运行此脚本文件。\n');
end