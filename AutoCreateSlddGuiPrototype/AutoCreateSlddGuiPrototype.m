function AutoCreateSlddGuiPrototype()
% AUTOCREATESLDDGUIPROTOTYPE AutoCreateSldd 的独立 GUI 原型
%
% 说明：
%   - 仅用于测试/验证阶段
%   - 不修改 main.m、Excel2Workspace.m、generateTemplate.m 的现有调用方式
%   - 通过独立 GUI 串联“选文件 -> 选工作表 -> 校验 -> 生成脚本 -> 导入字典”流程
%
% 用法：
%   AutoCreateSlddGuiPrototype

    app = struct();
    app.knownSheets = {'Signal', 'Parameter', 'Bus', 'BusElement'};

    app.fig = figure( ...
        'Name', 'AutoCreateSldd GUI 原型', ...
        'Tag', 'AutoCreateSlddGuiPrototypeFigure', ...
        'NumberTitle', 'off', ...
        'MenuBar', 'none', ...
        'ToolBar', 'none', ...
        'Resize', 'off', ...
        'Position', [40, 40, 1380, 820], ...
        'Color', [0.94, 0.94, 0.94]);

    app.filePanel = uipanel(app.fig, ...
        'Title', '文件与输出', ...
        'FontWeight', 'bold', ...
        'Position', [0.015, 0.79, 0.97, 0.19]);

    uicontrol(app.filePanel, 'Style', 'text', 'Position', [15, 92, 90, 22], ...
        'String', 'Excel 文件', 'HorizontalAlignment', 'left', 'FontWeight', 'bold');
    app.excelPathEdit = uicontrol(app.filePanel, 'Style', 'edit', 'Tag', 'excelPathEdit', ...
        'Position', [105, 92, 980, 26], 'HorizontalAlignment', 'left', 'BackgroundColor', 'white');
    app.browseExcelBtn = uicontrol(app.filePanel, 'Style', 'pushbutton', 'Position', [1100, 90, 90, 30], ...
        'String', '浏览...', 'Callback', @onBrowseExcel);
    app.loadSheetsBtn = uicontrol(app.filePanel, 'Style', 'pushbutton', 'Tag', 'loadSheetsBtn', ...
        'Position', [1200, 90, 120, 30], 'String', '载入工作表', 'Callback', @onLoadSheets);

    uicontrol(app.filePanel, 'Style', 'text', 'Position', [15, 54, 90, 22], ...
        'String', '输出脚本', 'HorizontalAlignment', 'left', 'FontWeight', 'bold');
    app.outputPathEdit = uicontrol(app.filePanel, 'Style', 'edit', 'Tag', 'outputPathEdit', ...
        'Position', [105, 54, 980, 26], 'HorizontalAlignment', 'left', 'BackgroundColor', 'white');
    app.browseOutputBtn = uicontrol(app.filePanel, 'Style', 'pushbutton', 'Position', [1100, 52, 90, 30], ...
        'String', '另存为...', 'Callback', @onBrowseOutput);

    uicontrol(app.filePanel, 'Style', 'text', 'Position', [15, 16, 90, 22], ...
        'String', '数据字典', 'HorizontalAlignment', 'left', 'FontWeight', 'bold');
    app.slddPathEdit = uicontrol(app.filePanel, 'Style', 'edit', 'Tag', 'slddPathEdit', ...
        'Position', [105, 16, 980, 26], 'HorizontalAlignment', 'left', 'BackgroundColor', 'white');
    app.browseSlddBtn = uicontrol(app.filePanel, 'Style', 'pushbutton', 'Position', [1100, 14, 90, 30], ...
        'String', '另存为...', 'Callback', @onBrowseSldd);
    app.generateTemplateBtn = uicontrol(app.filePanel, 'Style', 'pushbutton', 'Position', [1200, 14, 120, 68], ...
        'String', '生成模板', 'Callback', @onGenerateTemplate, 'FontWeight', 'bold');

    app.sheetPanel = uipanel(app.fig, ...
        'Title', '工作表选择', ...
        'FontWeight', 'bold', ...
        'Position', [0.015, 0.43, 0.24, 0.34]);

    app.sheetList = uicontrol(app.sheetPanel, 'Style', 'listbox', 'Tag', 'sheetList', ...
        'Position', [12, 72, 300, 185], 'Min', 0, 'Max', 10, 'BackgroundColor', 'white', ...
        'String', {}, 'Callback', @onSheetSelectionChanged);
    app.selectRecommendedBtn = uicontrol(app.sheetPanel, 'Style', 'pushbutton', 'Position', [12, 32, 94, 30], ...
        'String', '标准勾选', 'Callback', @onSelectRecommended);
    app.selectAllBtn = uicontrol(app.sheetPanel, 'Style', 'pushbutton', 'Position', [115, 32, 94, 30], ...
        'String', '全选', 'Callback', @onSelectAll);
    app.clearSelectionBtn = uicontrol(app.sheetPanel, 'Style', 'pushbutton', 'Position', [218, 32, 94, 30], ...
        'String', '清空', 'Callback', @onClearSelection);

    app.paramPanel = uipanel(app.fig, ...
        'Title', '参数区', ...
        'FontWeight', 'bold', ...
        'Position', [0.015, 0.21, 0.24, 0.20]);

    app.validateBeforeRunChk = uicontrol(app.paramPanel, 'Style', 'checkbox', 'Position', [12, 108, 210, 24], ...
        'String', '执行前先做数据校验', 'Value', 1, 'BackgroundColor', [0.94, 0.94, 0.94]);
    app.confirmBeforeSaveChk = uicontrol(app.paramPanel, 'Style', 'checkbox', 'Position', [12, 82, 210, 24], ...
        'String', '写入字典前弹确认框', 'Value', 1, 'BackgroundColor', [0.94, 0.94, 0.94]);
    app.onlySaveGeneratedChk = uicontrol(app.paramPanel, 'Style', 'checkbox', 'Position', [12, 56, 280, 24], ...
        'String', '仅写入本次生成的变量（推荐）', 'Value', 1, 'BackgroundColor', [0.94, 0.94, 0.94]);

    uicontrol(app.paramPanel, 'Style', 'text', 'Position', [12, 22, 120, 22], ...
        'String', '工作表预览行数', 'HorizontalAlignment', 'left');
    app.previewRowsEdit = uicontrol(app.paramPanel, 'Style', 'edit', 'Position', [110, 22, 60, 24], ...
        'String', '12', 'BackgroundColor', 'white', 'HorizontalAlignment', 'center');
    uicontrol(app.paramPanel, 'Style', 'text', 'Position', [185, 22, 118, 22], ...
        'String', '错误定位会围绕该行预览', 'HorizontalAlignment', 'left', 'ForegroundColor', [0.4, 0.4, 0.4]);

    app.actionPanel = uipanel(app.fig, ...
        'Title', '操作', ...
        'FontWeight', 'bold', ...
        'Position', [0.015, 0.04, 0.24, 0.15]);

    app.validateBtn = uicontrol(app.actionPanel, 'Style', 'pushbutton', 'Tag', 'validateBtn', ...
        'Position', [12, 62, 145, 34], 'String', '校验所选工作表', 'Callback', @onValidate, 'FontWeight', 'bold');
    app.previewSheetBtn = uicontrol(app.actionPanel, 'Style', 'pushbutton', 'Position', [167, 62, 145, 34], ...
        'String', '预览所选工作表', 'Callback', @onPreviewCurrentSheet, 'FontWeight', 'bold');
    app.generateBtn = uicontrol(app.actionPanel, 'Style', 'pushbutton', 'Tag', 'generateBtn', ...
        'Position', [12, 16, 145, 34], 'String', '仅生成 objects', 'Callback', @onGenerateOnly, 'FontWeight', 'bold');
    app.importBtn = uicontrol(app.actionPanel, 'Style', 'pushbutton', 'Tag', 'importBtn', ...
        'Position', [167, 16, 145, 34], 'String', '完整导入到 SLDD', 'Callback', @onFullImport, ...
        'FontWeight', 'bold', 'ForegroundColor', [0.1, 0.1, 0.1], 'BackgroundColor', [0.82, 0.9, 0.82]);

    app.tabGroup = uitabgroup(app.fig, 'Position', [0.27, 0.20, 0.715, 0.57]);

    app.sheetPreviewTab = uitab(app.tabGroup, 'Title', '工作表预览');
    uicontrol(app.sheetPreviewTab, 'Style', 'text', 'Position', [12, 326, 80, 22], ...
        'String', '预览工作表', 'HorizontalAlignment', 'left', 'FontWeight', 'bold');
    app.previewSheetPopup = uicontrol(app.sheetPreviewTab, 'Style', 'popupmenu', 'Position', [92, 326, 220, 24], ...
        'BackgroundColor', 'white', 'String', {'(未载入)'}, 'Callback', @onPreviewCurrentSheet);
    app.previewSummaryText = uicontrol(app.sheetPreviewTab, 'Style', 'text', 'Position', [330, 320, 590, 32], ...
        'String', {'尚未预览工作表'}, 'HorizontalAlignment', 'left', 'BackgroundColor', [0.94, 0.94, 0.94]);
    app.sheetPreviewTable = uitable(app.sheetPreviewTab, 'Position', [12, 12, 928, 302], ...
        'Data', cell(0, 1), 'ColumnName', {'预览结果'}, 'RowName', []);

    app.resultPreviewTab = uitab(app.tabGroup, 'Title', '结果预览');
    app.resultSummaryText = uicontrol(app.resultPreviewTab, 'Style', 'text', 'Position', [12, 312, 928, 40], ...
        'String', {'尚未生成结果'}, 'HorizontalAlignment', 'left', 'BackgroundColor', [0.94, 0.94, 0.94]);
    app.resultTable = uitable(app.resultPreviewTab, 'Position', [12, 12, 928, 292], ...
        'Data', cell(0, 4), 'ColumnName', {'变量名', '对象类型', '构造器', '状态'}, 'RowName', []);

    app.errorTab = uitab(app.tabGroup, 'Title', '错误定位');
    app.errorHintText = uicontrol(app.errorTab, 'Style', 'text', 'Position', [12, 322, 928, 24], ...
        'String', {'校验问题会拆成 sheet / 行 / 列；点击某一行可自动定位到对应工作表附近。'}, ...
        'HorizontalAlignment', 'left', 'BackgroundColor', [0.94, 0.94, 0.94]);
    app.errorTable = uitable(app.errorTab, 'Position', [12, 12, 928, 302], ...
        'Data', cell(0, 5), 'ColumnName', {'类型', 'Sheet', '行', '列', '说明'}, 'RowName', [], ...
        'CellSelectionCallback', @onErrorCellSelected);

    app.logPanel = uipanel(app.fig, ...
        'Title', '运行日志', ...
        'FontWeight', 'bold', ...
        'Position', [0.27, 0.04, 0.715, 0.14]);
    app.logBox = uicontrol(app.logPanel, 'Style', 'edit', 'Position', [12, 12, 840, 90], ...
        'Max', 200, 'Min', 0, 'HorizontalAlignment', 'left', 'BackgroundColor', 'white');
    app.clearLogBtn = uicontrol(app.logPanel, 'Style', 'pushbutton', 'Position', [865, 52, 75, 28], ...
        'String', '清空日志', 'Callback', @onClearLog);
    app.clearIssueBtn = uicontrol(app.logPanel, 'Style', 'pushbutton', 'Position', [865, 18, 75, 28], ...
        'String', '清空错误', 'Callback', @onClearIssues);

    app.statusText = uicontrol(app.fig, 'Style', 'text', 'Position', [20, 8, 1320, 22], ...
        'String', '就绪', 'HorizontalAlignment', 'left', 'BackgroundColor', [0.94, 0.94, 0.94], ...
        'ForegroundColor', [0.2, 0.2, 0.2], 'FontWeight', 'bold');

    setappdata(app.fig, 'AllSheets', {});
    setappdata(app.fig, 'ParsedIssues', struct('Type', {}, 'Sheet', {}, 'Row', {}, 'Column', {}, 'Message', {}));
    setappdata(app.fig, 'LastGeneratedInfo', struct('Name', {}, 'Constructor', {}, 'Type', {}));

    appendLog('GUI 原型已启动。当前版本会按“本次生成结果”同步数据字典，避免未选工作表的旧条目残留。');

    resetSheetPreview('尚未载入 Excel 工作表。');
    resetResultPreview('尚未生成结果。');
    resetIssueTable('尚未执行校验。');
    tryLoadDefaultWorkbook();

    function onBrowseExcel(~, ~)
        startDir = getSuggestedFolder();
        [fileName, filePath] = uigetfile({'*.xlsx;*.xls', 'Excel 文件 (*.xlsx, *.xls)'}, '选择数据字典 Excel', startDir);
        if isequal(fileName, 0)
            appendLog('已取消选择 Excel 文件。');
            return;
        end
        fullpath = fullfile(filePath, fileName);
        set(app.excelPathEdit, 'String', fullpath);
        loadWorkbook(fullpath);
    end

    function onLoadSheets(~, ~)
        excelPath = getRequiredFilePath(app.excelPathEdit, '请先选择 Excel 文件。');
        loadWorkbook(excelPath);
    end

    function loadWorkbook(excelPath)
        try
            if ~exist(excelPath, 'file')
                error('文件不存在：%s', excelPath);
            end
            refreshDerivedPaths(excelPath);
            clearPreviewAndIssues();
            loadSheetsFromExcel();
            appendLog(['当前 Excel：', excelPath]);
        catch ME
            appendLog(['载入 Excel 失败：', ME.message]);
            setStatus('载入 Excel 失败。', true);
            showErrorDialog(ME.message, '载入失败');
        end
    end

    function onBrowseOutput(~, ~)
        basePath = strtrim(get(app.outputPathEdit, 'String'));
        if isempty(basePath)
            excelPath = strtrim(get(app.excelPathEdit, 'String'));
            if isempty(excelPath)
                basePath = fullfile(pwd, 'MCU_Template_objects.m');
            else
                [folder, baseName] = fileparts(excelPath);
                basePath = fullfile(folder, [baseName, '_objects.m']);
            end
        end
        [fileName, filePath] = uiputfile({'*.m', 'MATLAB 脚本 (*.m)'}, '选择输出脚本位置', basePath);
        if isequal(fileName, 0)
            return;
        end
        set(app.outputPathEdit, 'String', fullfile(filePath, fileName));
        appendLog('已更新输出脚本路径。');
    end

    function onBrowseSldd(~, ~)
        basePath = strtrim(get(app.slddPathEdit, 'String'));
        if isempty(basePath)
            excelPath = strtrim(get(app.excelPathEdit, 'String'));
            if isempty(excelPath)
                basePath = fullfile(pwd, 'MCU_Template.sldd');
            else
                [folder, baseName] = fileparts(excelPath);
                basePath = fullfile(folder, [baseName, '.sldd']);
            end
        end
        [fileName, filePath] = uiputfile({'*.sldd', 'Simulink Data Dictionary (*.sldd)'}, '选择数据字典输出位置', basePath);
        if isequal(fileName, 0)
            return;
        end
        set(app.slddPathEdit, 'String', fullfile(filePath, fileName));
        appendLog('已更新数据字典路径。');
    end

    function onGenerateTemplate(~, ~)
        excelPath = strtrim(get(app.excelPathEdit, 'String'));
        if isempty(excelPath)
            defaultTarget = fullfile(pwd, 'MCU_Template.xlsx');
        else
            [folder, ~, ~] = fileparts(excelPath);
            defaultTarget = fullfile(folder, 'MCU_Template.xlsx');
        end

        [fileName, filePath] = uiputfile({'*.xlsx', 'Excel 文件 (*.xlsx)'}, '生成模板到', defaultTarget);
        if isequal(fileName, 0)
            appendLog('已取消生成模板。');
            return;
        end

        try
            templatePath = generateTemplate(filePath, fileName);
            set(app.excelPathEdit, 'String', templatePath);
            appendLog(['模板已生成：', templatePath]);
            loadWorkbook(templatePath);
            setStatus('模板生成成功。');
            showInfoDialog('模板已生成并自动载入 GUI。', '生成模板');
        catch ME
            appendLog(['模板生成失败：', ME.message]);
            setStatus('模板生成失败。', true);
            showErrorDialog(ME.message, '生成模板失败');
        end
    end

    function onSelectRecommended(~, ~)
        allSheets = getAllSheets();
        if isempty(allSheets)
            appendLog('还没有载入工作表，先选 Excel。');
            return;
        end
        selectIdx = find(ismember(allSheets, app.knownSheets));
        if isempty(selectIdx)
            selectIdx = 1:numel(allSheets);
        end
        set(app.sheetList, 'Value', selectIdx);
        onSheetSelectionChanged();
        appendLog('已按标准工作表预选。');
    end

    function onSelectAll(~, ~)
        allSheets = getAllSheets();
        if isempty(allSheets)
            return;
        end
        set(app.sheetList, 'Value', 1:numel(allSheets));
        onSheetSelectionChanged();
        appendLog('已全选所有工作表。');
    end

    function onClearSelection(~, ~)
        set(app.sheetList, 'Value', []);
        updatePreviewSheetChoices({});
        resetSheetPreview('尚未选择工作表。');
        setStatus('已清空工作表选择。');
        appendLog('已清空工作表选择。');
    end

    function onSheetSelectionChanged(~, ~)
        selectedSheets = safeGetSelectedSheets();
        updatePreviewSheetChoices(selectedSheets);
        if isempty(selectedSheets)
            resetSheetPreview('尚未选择工作表。');
            return;
        end
        previewSheetName = getCurrentPreviewSheetName();
        if isempty(previewSheetName)
            previewSheetName = selectedSheets{1};
        end
        showSheetPreview(previewSheetName, []);
    end

    function onPreviewCurrentSheet(~, ~)
        previewSheetName = getCurrentPreviewSheetName();
        if isempty(previewSheetName)
            resetSheetPreview('没有可预览的工作表。');
            return;
        end
        showSheetPreview(previewSheetName, []);
    end

    function onValidate(~, ~)
        try
            [isValid, issues, selectedSheets] = validateSelection();
            appendLog(['本次校验工作表：', strjoin(selectedSheets, ', ')]);
            if isValid
                resetIssueTable('校验通过，未发现结构性问题。');
                appendLog('校验通过，未发现结构性问题。');
                setStatus('校验通过。');
                showInfoDialog('校验通过，未发现结构性问题。', '校验结果');
            else
                parsedIssues = updateIssueTable(issues);
                appendLog(sprintf('校验失败，共发现 %d 条问题。', numel(parsedIssues)));
                for i = 1:numel(issues)
                    appendLog(['  - ', issues{i}]);
                end
                if ~isempty(parsedIssues)
                    jumpToIssue(parsedIssues(1));
                end
                setStatus('校验失败。', true);
                showErrorDialog(strjoin(issues, newline), '校验失败');
            end
        catch ME
            appendLog(['校验执行失败：', ME.message]);
            setStatus('校验执行失败。', true);
            showErrorDialog(ME.message, '校验失败');
        end
    end

    function onGenerateOnly(~, ~)
        try
            [outputPath, selectedSheets, generatedInfo] = doGenerateObjects();
            appendLog(['objects 脚本已生成：', outputPath]);
            setStatus('objects 脚本生成完成。');
            showGeneratedResultPreview(generatedInfo, selectedSheets, false, '已生成 objects 脚本，尚未执行到基础工作区。');
            showInfoDialog('objects 脚本已生成，可直接检查输出文件。', '生成完成');
        catch ME
            appendLog(['生成 objects 脚本失败：', ME.message]);
            setStatus('objects 脚本生成失败。', true);
            showErrorDialog(ME.message, '生成失败');
        end
    end

    function onFullImport(~, ~)
        try
            [outputPath, selectedSheets, generatedInfo] = doGenerateObjects();
            scriptContent = fileread(outputPath);
            try
                evalin('base', scriptContent);
            catch ME
                error('生成脚本执行失败，数据字典尚未写入，因此你现在看到的旧条目并不代表工作表选择失效。原始错误：%s', ME.message);
            end
            appendLog('已将生成脚本执行到 MATLAB 基础工作区。');

            slddPath = getRequiredFilePath(app.slddPathEdit, '请先填写数据字典输出路径。');
            confirmSave = logical(get(app.confirmBeforeSaveChk, 'Value'));
            onlySaveGenerated = logical(get(app.onlySaveGeneratedChk, 'Value'));
            syncToGeneratedScope = false;
            if onlySaveGenerated
                saveVarNames = {generatedInfo.Name};
                syncToGeneratedScope = true;
                appendLog(sprintf('本次将按当前生成结果同步数据字典，共 %d 个顶层对象；未出现在本次结果中的旧受管条目会被清理。', numel(saveVarNames)));
            else
                saveVarNames = {};
                appendLog('本次将写入基础工作区全部变量，不自动清理数据字典中的旧条目。');
            end

            saveVarsToDataDictionaryLocal(slddPath, saveVarNames, confirmSave, @appendLog, syncToGeneratedScope);
            showGeneratedResultPreview(generatedInfo, selectedSheets, true, '已执行脚本并完成数据字典写入。');
            setStatus('完整导入流程执行完成。');
            showInfoDialog('已完成 objects 生成、脚本执行和数据字典写入。', '导入完成');
        catch ME
            appendLog(['完整导入失败：', ME.message]);
            setStatus('完整导入失败。', true);
            showErrorDialog(ME.message, '导入失败');
        end
    end


    function onClearLog(~, ~)
        set(app.logBox, 'String', {});
        setStatus('日志已清空。');
    end

    function onClearIssues(~, ~)
        resetIssueTable('错误列表已清空。');
        setStatus('错误列表已清空。');
    end

    function onErrorCellSelected(~, event)
        if isempty(event.Indices)
            return;
        end
        rowIdx = event.Indices(1);
        parsedIssues = getappdata(app.fig, 'ParsedIssues');
        if isempty(parsedIssues) || rowIdx > numel(parsedIssues)
            return;
        end
        jumpToIssue(parsedIssues(rowIdx));
    end

    function [outputPath, selectedSheets, generatedInfo] = doGenerateObjects()
        excelPath = getRequiredFilePath(app.excelPathEdit, '请先选择 Excel 文件。');
        outputPath = getRequiredFilePath(app.outputPathEdit, '请先填写输出脚本路径。');
        selectedSheets = getSelectedSheets();
        allSheets = getAllSheets();

        if logical(get(app.validateBeforeRunChk, 'Value'))
            [isValid, issues] = validateDataDictionary(excelPath, selectedSheets);
            if ~isValid
                parsedIssues = updateIssueTable(issues);
                appendLog('执行前校验未通过，已停止生成。');
                for i = 1:numel(issues)
                    appendLog(['  - ', issues{i}]);
                end
                if ~isempty(parsedIssues)
                    jumpToIssue(parsedIssues(1));
                end
                error('执行前校验未通过，请先修正 Excel 数据。');
            end
            resetIssueTable('执行前校验通过。');
            appendLog('执行前校验通过。');
        end

        [folder, ~, ext] = fileparts(outputPath);
        if isempty(ext)
            outputPath = [outputPath, '.m'];
            set(app.outputPathEdit, 'String', outputPath);
            [folder, ~, ~] = fileparts(outputPath);
        end
        if ~isempty(folder) && ~exist(folder, 'dir')
            mkdir(folder);
        end

        [~, fileName, fileExt] = fileparts(excelPath);
        selectedIndices = find(ismember(allSheets, selectedSheets));
        if isempty(selectedIndices)
            error('没有可用于生成的工作表。');
        end

        appendLog(['准备生成 objects 脚本：', outputPath]);
        appendLog(['本次工作表：', strjoin(selectedSheets, ', ')]);
        runExcel2WorkspaceWithSelection(allSheets, selectedIndices, excelPath, outputPath, [fileName, fileExt]);

        generatedInfo = collectGeneratedObjectInfo(outputPath);
        generatedInfo = filterExportedObjectInfo(generatedInfo);
        setappdata(app.fig, 'LastGeneratedInfo', generatedInfo);

    end

    function [isValid, issues, selectedSheets] = validateSelection()
        excelPath = getRequiredFilePath(app.excelPathEdit, '请先选择 Excel 文件。');
        selectedSheets = getSelectedSheets();
        [isValid, issues] = validateDataDictionary(excelPath, selectedSheets);
    end

    function loadSheetsFromExcel()
        excelPath = getRequiredFilePath(app.excelPathEdit, '请先选择 Excel 文件。');
        if ~exist(excelPath, 'file')
            error('文件不存在：%s', excelPath);
        end

        allSheets = cellstr(sheetnames(excelPath));
        setappdata(app.fig, 'AllSheets', allSheets);
        set(app.sheetList, 'String', allSheets, 'Value', defaultSelection(allSheets));
        updatePreviewSheetChoices(allSheets(defaultSelection(allSheets)));
        appendLog(['已载入工作表：', strjoin(allSheets, ', ')]);
        setStatus('工作表载入完成。');
        if ~isempty(allSheets)
            defaultIdx = defaultSelection(allSheets, 1);
            showSheetPreview(allSheets{defaultIdx(1)}, []);
        else
            resetSheetPreview('Excel 中没有可用工作表。');
        end
    end

    function tryLoadDefaultWorkbook()
        candidate = findDefaultExcelInCurrentFolder();
        if isempty(candidate)
            return;
        end
        set(app.excelPathEdit, 'String', candidate);
        try
            loadWorkbook(candidate);
        catch ME
            appendLog(['自动载入默认 Excel 失败：', ME.message]);
        end
    end

    function allSheets = getAllSheets()
        allSheets = getappdata(app.fig, 'AllSheets');
        if isempty(allSheets)
            allSheets = {};
        end
    end

    function selectedSheets = getSelectedSheets()
        allSheets = getAllSheets();
        if isempty(allSheets)
            error('还没有载入工作表，请先点击“载入工作表”。');
        end
        selectedIdx = get(app.sheetList, 'Value');
        if isempty(selectedIdx)
            error('请至少选择一个工作表。');
        end
        selectedSheets = allSheets(selectedIdx);
    end

    function selectedSheets = safeGetSelectedSheets()
        try
            selectedSheets = getSelectedSheets();
        catch
            selectedSheets = {};
        end
    end

    function updatePreviewSheetChoices(candidateSheets)
        allSheets = getAllSheets();
        if nargin < 1 || isempty(candidateSheets)
            candidateSheets = allSheets;
        end
        if isempty(candidateSheets)
            set(app.previewSheetPopup, 'String', {'(未载入)'}, 'Value', 1);
            return;
        end

        currentSelection = getCurrentPreviewSheetName();
        popupStrings = candidateSheets;
        set(app.previewSheetPopup, 'String', popupStrings);
        idx = find(strcmp(popupStrings, currentSelection), 1);
        if isempty(idx)
            idx = 1;
        end
        set(app.previewSheetPopup, 'Value', idx);
    end

    function sheetName = getCurrentPreviewSheetName()
        popupItems = get(app.previewSheetPopup, 'String');
        if isempty(popupItems)
            sheetName = '';
            return;
        end
        if ischar(popupItems)
            popupItems = {popupItems};
        end
        popupValue = get(app.previewSheetPopup, 'Value');
        popupValue = min(max(1, popupValue), numel(popupItems));
        sheetName = popupItems{popupValue};
        if strcmp(sheetName, '(未载入)')
            sheetName = '';
        end
    end

    function showSheetPreview(sheetName, focusRow)
        excelPath = strtrim(get(app.excelPathEdit, 'String'));
        if isempty(excelPath) || isempty(sheetName)
            resetSheetPreview('没有可预览的工作表。');
            return;
        end

        previewRows = getPreviewRowCount();
        [sheetTable, readIssue] = readSheetTableAsStrings(excelPath, sheetName);
        if ~isempty(readIssue)
            resetSheetPreview(['读取工作表失败：', readIssue]);
            setStatus('工作表预览失败。', true);
            return;
        end

        [previewData, previewColumns, summaryLines] = convertTableToPreview(sheetTable, previewRows, focusRow);
        set(app.sheetPreviewTable, 'Data', previewData, 'ColumnName', previewColumns);
        set(app.previewSummaryText, 'String', summaryLines);
        set(app.tabGroup, 'SelectedTab', app.sheetPreviewTab);
        setStatus(['已预览工作表：', sheetName]);
    end

    function resetSheetPreview(message)
        set(app.sheetPreviewTable, 'Data', cell(0, 1), 'ColumnName', {'预览结果'});
        set(app.previewSummaryText, 'String', {message});
    end

    function showGeneratedResultPreview(generatedInfo, selectedSheets, includeWorkspaceStatus, stageMessage)
        if nargin < 4 || isempty(stageMessage)
            stageMessage = '结果已刷新。';
        end
        if nargin < 3
            includeWorkspaceStatus = false;
        end
        setappdata(app.fig, 'LastGeneratedInfo', generatedInfo);

        if isempty(generatedInfo)
            resetResultPreview('本次没有识别到可导入对象。');
            setStatus('本次没有识别到可导入对象。', true);
            return;
        end

        resultData = buildResultPreviewData(generatedInfo, includeWorkspaceStatus);
        summaryLines = buildGeneratedSummaryLines(generatedInfo, selectedSheets, stageMessage, includeWorkspaceStatus);
        set(app.resultTable, 'Data', resultData, 'ColumnName', {'变量名', '对象类型', '构造器', '状态'});
        set(app.resultSummaryText, 'String', summaryLines);
        set(app.tabGroup, 'SelectedTab', app.resultPreviewTab);
    end

    function resetResultPreview(message)
        set(app.resultTable, 'Data', cell(0, 4), 'ColumnName', {'变量名', '对象类型', '构造器', '状态'});
        set(app.resultSummaryText, 'String', {message});
    end

    function parsedIssues = updateIssueTable(issues)
        parsedIssues = parseValidationIssues(issues);
        setappdata(app.fig, 'ParsedIssues', parsedIssues);

        if isempty(parsedIssues)
            resetIssueTable('未解析到结构化错误信息。');
            return;
        end

        issueData = cell(numel(parsedIssues), 5);
        for i = 1:numel(parsedIssues)
            issueData{i, 1} = parsedIssues(i).Type;
            issueData{i, 2} = parsedIssues(i).Sheet;
            if isnan(parsedIssues(i).Row)
                issueData{i, 3} = '';
            else
                issueData{i, 3} = parsedIssues(i).Row;
            end
            issueData{i, 4} = parsedIssues(i).Column;
            issueData{i, 5} = parsedIssues(i).Message;
        end

        set(app.errorTable, 'Data', issueData, 'ColumnName', {'类型', 'Sheet', '行', '列', '说明'});
        set(app.errorHintText, 'String', {sprintf('共解析到 %d 条错误定位信息；点击某一行可自动跳到对应工作表附近。', numel(parsedIssues))});
        set(app.tabGroup, 'SelectedTab', app.errorTab);
    end

    function resetIssueTable(message)
        set(app.errorTable, 'Data', cell(0, 5), 'ColumnName', {'类型', 'Sheet', '行', '列', '说明'});
        set(app.errorHintText, 'String', {message});
        setappdata(app.fig, 'ParsedIssues', struct('Type', {}, 'Sheet', {}, 'Row', {}, 'Column', {}, 'Message', {}));
    end

    function jumpToIssue(issue)
        if isempty(issue.Sheet)
            set(app.tabGroup, 'SelectedTab', app.errorTab);
            return;
        end

        updatePreviewSheetChoices({issue.Sheet});
        set(app.previewSheetPopup, 'String', {issue.Sheet}, 'Value', 1);
        if isnan(issue.Row)
            showSheetPreview(issue.Sheet, []);
            setStatus(['已定位到工作表：', issue.Sheet], true);
        else
            showSheetPreview(issue.Sheet, issue.Row);
            setStatus(sprintf('已定位到 %s 第 %d 行附近。', issue.Sheet, issue.Row), true);
        end
    end

    function clearPreviewAndIssues()
        resetSheetPreview('尚未预览工作表。');
        resetResultPreview('尚未生成结果。');
        resetIssueTable('尚未执行校验。');
        setappdata(app.fig, 'LastGeneratedInfo', struct('Name', {}, 'Constructor', {}, 'Type', {}));
    end

    function previewRows = getPreviewRowCount()
        previewRows = str2double(strtrim(get(app.previewRowsEdit, 'String')));
        if isnan(previewRows) || previewRows < 1
            previewRows = 12;
            set(app.previewRowsEdit, 'String', '12');
        end
        previewRows = max(1, round(previewRows));
    end

    function refreshDerivedPaths(excelPath)
        if isempty(excelPath)
            return;
        end
        [folder, baseName, ~] = fileparts(excelPath);
        set(app.outputPathEdit, 'String', fullfile(folder, [baseName, '_objects.m']));
        set(app.slddPathEdit, 'String', fullfile(folder, [baseName, '.sldd']));
    end

    function appendLog(message)
        timestamp = char(string(datetime('now', 'Format', 'HH:mm:ss')));
        newLine = [timestamp, '  ', message];
        currentText = get(app.logBox, 'String');
        if isempty(currentText)
            currentText = {newLine};
        elseif ischar(currentText)
            currentText = {currentText; newLine};
        else
            currentText{end + 1, 1} = newLine;
        end
        set(app.logBox, 'String', currentText);
        drawnow;
    end

    function setStatus(message, isError)
        if nargin < 2
            isError = false;
        end
        set(app.statusText, 'String', message);
        if isError
            set(app.statusText, 'ForegroundColor', [0.75, 0.1, 0.1]);
        else
            set(app.statusText, 'ForegroundColor', [0.2, 0.2, 0.2]);
        end
        drawnow;
    end

    function fullpath = getRequiredFilePath(handleObj, errMsg)
        fullpath = strtrim(get(handleObj, 'String'));
        if isempty(fullpath)
            error(errMsg);
        end
    end

    function folder = getSuggestedFolder()
        excelPath = strtrim(get(app.excelPathEdit, 'String'));
        if ~isempty(excelPath)
            folder = fileparts(excelPath);
            if exist(folder, 'dir')
                return;
            end
        end
        folder = pwd;
    end

    function showInfoDialog(message, titleText)
        if nargin < 2
            titleText = '提示';
        end
        if usejava('desktop')
            msgbox(message, titleText, 'modal');
        else
            appendLog([titleText, '：', message]);
        end
    end

    function showErrorDialog(message, titleText)
        if nargin < 2
            titleText = '错误';
        end
        if usejava('desktop')
            errordlg(message, titleText, 'modal');
        else
            appendLog([titleText, '：', message]);
        end
    end
end


function selection = defaultSelection(allSheets, fallbackValue)
    if nargin < 2
        fallbackValue = 1;
    end
    knownSheets = {'Signal', 'Parameter', 'Bus', 'BusElement'};
    selection = find(ismember(allSheets, knownSheets));
    if isempty(selection) && ~isempty(allSheets)
        selection = fallbackValue;
    end
end


function fullpath = findDefaultExcelInCurrentFolder()
    excelFiles = dir(fullfile(pwd, '*.xlsx'));
    if isempty(excelFiles)
        excelFiles = dir(fullfile(pwd, '*.xls'));
    end

    if isempty(excelFiles)
        fullpath = '';
    else
        fullpath = fullfile(pwd, excelFiles(1).name);
    end
end


function runExcel2WorkspaceWithSelection(allSheets, selectedIndices, excelPath, outputPath, fileName)
    if nargin < 5
        [~, nameOnly, ext] = fileparts(excelPath);
        fileName = [nameOnly, ext];
    end

    Excel2Workspace(allSheets, excelPath, outputPath, fileName, selectedIndices);
end



function saveVarsToDataDictionaryLocal(slddFullPath, targetVarNames, confirmSave, logFcn, syncManagedScope)
    if nargin < 2 || isempty(targetVarNames)
        targetVarNames = {};
    end
    if nargin < 3 || isempty(confirmSave)
        confirmSave = true;
    end
    if nargin < 4 || isempty(logFcn)
        logFcn = @(msg) fprintf('%s\n', msg);
    end
    if nargin < 5 || isempty(syncManagedScope)
        syncManagedScope = false;
    end


    [dictPath, dictName, fileExt] = fileparts(slddFullPath);
    if isempty(fileExt)
        slddFullPath = [slddFullPath, '.sldd'];
        fileExt = '.sldd';
    elseif ~strcmpi(fileExt, '.sldd')
        slddFullPath = fullfile(dictPath, [dictName, '.sldd']);
        fileExt = '.sldd';
    end

    if isempty(dictPath)
        dictPath = pwd;
        slddFullPath = fullfile(dictPath, [dictName, fileExt]);
    elseif ~exist(dictPath, 'dir')
        mkdir(dictPath);
    end

    try
        allVars = evalin('base', 'who');
    catch ME
        error('无法访问基础工作区：%s', ME.message);
    end

    excludeVars = {'ans'};
    allVars = setdiff(allVars, excludeVars, 'stable');

    if isempty(targetVarNames)
        varNames = allVars;
        logFcn(sprintf('准备写入基础工作区全部变量，共 %d 个。', numel(varNames)));
    else
        targetVarNames = cellstr(string(targetVarNames));
        targetVarNames = setdiff(targetVarNames, excludeVars, 'stable');
        varNames = targetVarNames(ismember(targetVarNames, allVars));
        missingNames = targetVarNames(~ismember(targetVarNames, allVars));
        if ~isempty(missingNames)
            logFcn(sprintf('以下变量未在基础工作区找到，已跳过：%s', strjoin(missingNames, ', ')));
        end
        logFcn(sprintf('准备仅写入本次生成变量，共 %d 个。', numel(varNames)));
    end

    if isempty(varNames)
        error('没有可写入数据字典的变量。');
    end

    if confirmSave
        previewCount = min(numel(varNames), 8);
        previewText = strjoin(varNames(1:previewCount), newline);
        if numel(varNames) > previewCount
            previewText = sprintf('%s\n... 还有 %d 个变量', previewText, numel(varNames) - previewCount);
        end
        if usejava('desktop')
            choice = questdlg(sprintf('即将写入 %d 个变量到数据字典：\n\n%s', numel(varNames), previewText), ...
                '确认写入数据字典', '继续', '取消', '继续');
            if ~strcmp(choice, '继续')
                logFcn('用户取消了数据字典写入。');
                return;
            end
        else
            logFcn('当前环境不支持交互确认，已按“继续”执行写入。');
        end
    end

    try
        if exist(slddFullPath, 'file')
            dictObj = Simulink.data.dictionary.open(slddFullPath);
            logFcn(['已打开现有数据字典：', slddFullPath]);
        else
            dictObj = Simulink.data.dictionary.create(slddFullPath);
            logFcn(['已创建新数据字典：', slddFullPath]);
        end
    catch ME
        error('无法打开或创建数据字典：%s', ME.message);
    end

    closeGuard = onCleanup(@() safeCloseDictionary(dictObj)); %#ok<NASGU>

    try
        dataSect = getSection(dictObj, 'Design Data');
    catch ME
        error('无法获取 Design Data 段：%s', ME.message);
    end

    removedCount = 0;
    if syncManagedScope && ~isempty(targetVarNames)
        [removedCount, removedNames] = removeManagedEntriesOutsideTarget(dataSect, targetVarNames, logFcn);
        if ~isempty(removedNames)
            logFcn(sprintf('已清理不在本次结果中的旧条目：%s', strjoin(removedNames, ', ')));
        else
            logFcn('本次无需清理旧条目。');
        end
    end

    addedCount = 0;
    updatedCount = 0;
    unchangedCount = 0;
    errorCount = 0;


    for i = 1:numel(varNames)
        varName = varNames{i};
        try
            varValue = evalin('base', varName);
            existingEntries = find(dataSect, 'Name', varName);
            if ~isempty(existingEntries)
                entryObj = existingEntries(1);
                currentValue = getValue(entryObj);
                if isequal(varValue, currentValue)
                    unchangedCount = unchangedCount + 1;
                    continue;
                end
                setValue(entryObj, varValue);
                updatedCount = updatedCount + 1;
                logFcn(['已更新：', varName]);
            else
                addEntry(dataSect, varName, varValue);
                addedCount = addedCount + 1;
                logFcn(['已添加：', varName]);
            end
        catch ME
            errorCount = errorCount + 1;
            logFcn(['保存失败：', varName, ' -> ', ME.message]);
        end
    end

    try
        saveChanges(dictObj);
        logFcn('数据字典更改已保存。');
    catch ME
        error('保存数据字典失败：%s', ME.message);
    end

    logFcn(sprintf('写入完成：清理 %d，新增 %d，更新 %d，跳过 %d，失败 %d。', ...
        removedCount, addedCount, updatedCount, unchangedCount, errorCount));


    try
        verifyDict = Simulink.data.dictionary.open(slddFullPath);
        verifyGuard = onCleanup(@() safeCloseDictionary(verifyDict)); %#ok<NASGU>
        verifySect = getSection(verifyDict, 'Design Data');
        dictEntries = find(verifySect);
        logFcn(sprintf('验证完成：当前数据字典共有 %d 个条目。', numel(dictEntries)));
    catch ME
        logFcn(['验证数据字典时出错：', ME.message]);
    end
end


function [removedCount, removedNames] = removeManagedEntriesOutsideTarget(dataSect, targetVarNames, logFcn)
    if nargin < 3 || isempty(logFcn)
        logFcn = @(msg) fprintf('%s\n', msg);
    end

    removedCount = 0;
    removedNames = {};
    targetVarNames = cellstr(string(targetVarNames));
    dictEntries = find(dataSect);

    for i = 1:numel(dictEntries)
        entryName = dictEntries(i).Name;
        if ismember(entryName, targetVarNames)
            continue;
        end

        try
            entryValue = getValue(dictEntries(i));
        catch ME
            logFcn(['读取现有条目失败，已跳过：', entryName, ' -> ', ME.message]);
            continue;
        end

        if ~isManagedDictionaryEntry(entryName, entryValue)
            continue;
        end

        deleteEntry(dataSect, entryName);
        removedCount = removedCount + 1;
        removedNames{end + 1} = entryName; %#ok<AGROW>
    end
end


function tf = isManagedDictionaryEntry(entryName, entryValue)
    managedTypes = {'Signal', 'Parameter', 'Bus', 'BusElement'};
    tf = ismember(getValueTypeSuffix(entryValue), managedTypes);
    if tf
        return;
    end

    entryName = char(string(entryName));
    tf = numel(entryName) >= 9 && strcmp(entryName(end-8:end), '_elements');
end


function typeName = getValueTypeSuffix(entryValue)
    valueClass = class(entryValue);
    dotIdx = find(valueClass == '.', 1, 'last');
    if isempty(dotIdx)
        typeName = valueClass;
    else
        typeName = valueClass(dotIdx + 1:end);
    end
end


function safeCloseDictionary(dictObj)

    try
        close(dictObj);
    catch
    end
end


function generatedInfo = collectGeneratedObjectInfo(outputPath)
    scriptContent = fileread(outputPath);
    lines = regexp(scriptContent, '\r\n|\n|\r', 'split');
    generatedInfo = struct('Name', {}, 'Constructor', {}, 'Type', {});
    seenNames = {};

    for i = 1:numel(lines)
        tokens = regexp(lines{i}, '^\s*([A-Za-z]\w*)\s*=\s*([A-Za-z]\w*(?:\.[A-Za-z]\w*)*)\s*;\s*$', 'tokens', 'once');
        if isempty(tokens)
            continue;
        end

        varName = tokens{1};
        constructor = tokens{2};
        if any(strcmp(seenNames, varName))
            continue;
        end

        seenNames{end + 1} = varName; %#ok<AGROW>
        generatedInfo(end + 1).Name = varName; %#ok<AGROW>
        generatedInfo(end).Constructor = constructor;
        generatedInfo(end).Type = inferGeneratedType(constructor);
    end
end


function generatedInfo = filterExportedObjectInfo(generatedInfo)
    if isempty(generatedInfo)
        return;
    end

    keepMask = true(1, numel(generatedInfo));
    for i = 1:numel(generatedInfo)
        keepMask(i) = ~isGeneratedHelperVariable(generatedInfo(i));
    end
    generatedInfo = generatedInfo(keepMask);
end


function tf = isGeneratedHelperVariable(info)
    tf = strcmp(info.Type, 'BusElement');
    if tf
        return;
    end

    nameText = char(string(info.Name));
    if numel(nameText) >= 9 && strcmp(nameText(end-8:end), '_elements')
        tf = true;
    end
end


function typeName = inferGeneratedType(constructor)
    dotIdx = find(constructor == '.', 1, 'last');
    if isempty(dotIdx)
        typeName = constructor;
    else
        typeName = constructor(dotIdx + 1:end);
    end
end


function resultData = buildResultPreviewData(generatedInfo, includeWorkspaceStatus)

    resultData = cell(numel(generatedInfo), 4);
    for i = 1:numel(generatedInfo)
        resultData{i, 1} = generatedInfo(i).Name;
        resultData{i, 2} = generatedInfo(i).Type;
        resultData{i, 3} = generatedInfo(i).Constructor;
        if includeWorkspaceStatus
            resultData{i, 4} = describeWorkspaceVariable(generatedInfo(i).Name);
        else
            resultData{i, 4} = '待执行';
        end
    end
end


function summaryLines = buildGeneratedSummaryLines(generatedInfo, selectedSheets, stageMessage, includeWorkspaceStatus)
    typeList = {generatedInfo.Type};
    uniqueTypes = unique(typeList, 'stable');
    typeSummary = cell(1, numel(uniqueTypes));
    for i = 1:numel(uniqueTypes)
        typeSummary{i} = sprintf('%s=%d', uniqueTypes{i}, sum(strcmp(typeList, uniqueTypes{i})));
    end

    if isempty(typeSummary)
        typeText = '无可识别对象';
    else
        typeText = strjoin(typeSummary, '，');
    end

    selectedText = strjoin(selectedSheets, ', ');
    if includeWorkspaceStatus
        stageSuffix = '状态列显示当前基础工作区中的实际类。';
    else
        stageSuffix = '状态列为待执行，便于先检查生成结果。';
    end

    summaryLines = { ...
        stageMessage, ...
        sprintf('工作表：%s | 识别对象数：%d | 类型分布：%s', selectedText, numel(generatedInfo), typeText), ...
        stageSuffix};
end


function statusText = describeWorkspaceVariable(varName)
    try
        allVars = evalin('base', 'who');
        if ~ismember(varName, allVars)
            statusText = '未加载到基础工作区';
            return;
        end
        varValue = evalin('base', varName);
        statusText = class(varValue);
    catch ME
        statusText = ['读取失败：', ME.message];
    end
end


function [sheetTable, readIssue] = readSheetTableAsStrings(excelPath, sheetName)
    try
        opts = detectImportOptions(excelPath, 'Sheet', sheetName);
        opts = setvartype(opts, opts.VariableNames, 'string');
        sheetTable = readtable(excelPath, opts);
        readIssue = '';
    catch ME
        sheetTable = table();
        readIssue = ME.message;
    end
end


function [previewData, previewColumns, summaryLines] = convertTableToPreview(sheetTable, previewRows, focusRow)
    previewColumns = [{'ExcelRow'}, sheetTable.Properties.VariableNames];
    totalRows = height(sheetTable);

    if totalRows == 0
        previewData = cell(0, numel(previewColumns));
        summaryLines = {'该工作表没有数据行。'};
        return;
    end

    excelRowNumbers = (2:(totalRows + 1))';
    if nargin < 3 || isempty(focusRow) || isnan(focusRow)
        startIdx = 1;
        endIdx = min(totalRows, previewRows);
        focusText = sprintf('显示前 %d 行数据。', endIdx - startIdx + 1);
    else
        nearestIdx = find(excelRowNumbers >= focusRow, 1, 'first');
        if isempty(nearestIdx)
            nearestIdx = totalRows;
        end
        halfWindow = floor(previewRows / 2);
        startIdx = max(1, nearestIdx - halfWindow);
        endIdx = min(totalRows, startIdx + previewRows - 1);
        startIdx = max(1, endIdx - previewRows + 1);
        focusText = sprintf('已围绕 Excel 第 %d 行定位，当前显示第 %d-%d 行数据。', focusRow, excelRowNumbers(startIdx), excelRowNumbers(endIdx));
    end

    previewSlice = sheetTable(startIdx:endIdx, :);
    previewCells = table2cell(previewSlice);
    previewCells = sanitizeUitableCells(previewCells);
    previewData = cell(size(previewCells, 1), size(previewCells, 2) + 1);
    previewData(:, 1) = num2cell(excelRowNumbers(startIdx:endIdx));
    previewData(:, 2:end) = previewCells;

    summaryLines = { ...
        sprintf('总数据行：%d，当前预览：%d 行。', totalRows, size(previewData, 1)), ...
        focusText};
end


function cellData = sanitizeUitableCells(cellData)
    for row = 1:size(cellData, 1)
        for col = 1:size(cellData, 2)
            value = cellData{row, col};
            if isstring(value)
                if ismissing(value)
                    cellData{row, col} = '';
                else
                    cellData{row, col} = char(value);
                end
            elseif ischar(value) || isnumeric(value) || islogical(value)
                cellData{row, col} = value;
            elseif isempty(value)
                cellData{row, col} = '';
            else
                cellData{row, col} = char(string(value));
            end
        end
    end
end


function parsedIssues = parseValidationIssues(issues)

    parsedIssues = struct('Type', {}, 'Sheet', {}, 'Row', {}, 'Column', {}, 'Message', {});

    for i = 1:numel(issues)
        issueText = char(string(issues{i}));

        tokens = regexp(issueText, '^工作表 (.+?) 的必填列 (.+?) 在第 ([0-9,\s]+) 行为空。$', 'tokens', 'once');
        if ~isempty(tokens)
            rowList = parseRowTokens(tokens{3});
            for row = rowList
                parsedIssues(end + 1) = makeParsedIssue('必填为空', tokens{1}, row, tokens{2}, issueText); %#ok<AGROW>
            end
            continue;
        end

        tokens = regexp(issueText, '^工作表 (.+?) 缺少必要列：(.+)$', 'tokens', 'once');
        if ~isempty(tokens)
            parsedIssues(end + 1) = makeParsedIssue('缺少列', tokens{1}, NaN, tokens{2}, issueText); %#ok<AGROW>
            continue;
        end

        tokens = regexp(issueText, '^变量名 (.+?) 重复出现：(.+)。$', 'tokens', 'once');
        if ~isempty(tokens)
            locationTokens = regexp(tokens{2}, '(Signal|Parameter) 第 (\d+) 行', 'tokens');
            if isempty(locationTokens)
                parsedIssues(end + 1) = makeParsedIssue('变量重复', '', NaN, tokens{1}, issueText); %#ok<AGROW>
            else
                for j = 1:numel(locationTokens)
                    parsedIssues(end + 1) = makeParsedIssue('变量重复', locationTokens{j}{1}, str2double(locationTokens{j}{2}), tokens{1}, issueText); %#ok<AGROW>
                end
            end
            continue;
        end

        tokens = regexp(issueText, '^Bus 名称 (.+?) 在 Bus 工作表中重复出现：第 ([0-9,\s]+) 行。$', 'tokens', 'once');
        if ~isempty(tokens)
            rowList = parseRowTokens(tokens{2});
            for row = rowList
                parsedIssues(end + 1) = makeParsedIssue('Bus 重复', 'Bus', row, 'BusName', issueText); %#ok<AGROW>
            end
            continue;
        end

        tokens = regexp(issueText, '^(Signal|Parameter) 的 DataType 引用了未在 Bus 表中定义的 Bus (.+?)：第 ([0-9,\s]+) 行。$', 'tokens', 'once');
        if ~isempty(tokens)
            rowList = parseRowTokens(tokens{3});
            for row = rowList
                parsedIssues(end + 1) = makeParsedIssue('Bus 引用缺失', tokens{1}, row, 'DataType', issueText); %#ok<AGROW>
            end
            continue;
        end

        tokens = regexp(issueText, '^(Signal|Parameter) 的 DataType 引用了 Bus (.+?)，但未同时选择 Bus：第 ([0-9,\s]+) 行。$', 'tokens', 'once');
        if ~isempty(tokens)
            rowList = parseRowTokens(tokens{3});
            for row = rowList
                parsedIssues(end + 1) = makeParsedIssue('选择冲突', tokens{1}, row, 'DataType', issueText); %#ok<AGROW>
            end
            continue;
        end

        tokens = regexp(issueText, '^BusElement 引用了未在 Bus 表中定义的 Bus (.+?)：第 ([0-9,\s]+) 行。$', 'tokens', 'once');
        if ~isempty(tokens)
            rowList = parseRowTokens(tokens{2});
            for row = rowList
                parsedIssues(end + 1) = makeParsedIssue('Bus 引用缺失', 'BusElement', row, 'BusName', issueText); %#ok<AGROW>
            end
            continue;
        end

        if contains(issueText, '已选择 BusElement，但未同时选择 Bus')
            parsedIssues(end + 1) = makeParsedIssue('选择冲突', 'BusElement', NaN, '', issueText); %#ok<AGROW>
            continue;
        end


        parsedIssues(end + 1) = makeParsedIssue('通用问题', '', NaN, '', issueText); %#ok<AGROW>
    end
end


function rowList = parseRowTokens(rowText)
    rowTokens = regexp(rowText, '\d+', 'match');
    if isempty(rowTokens)
        rowList = NaN;
    else
        rowList = str2double(rowTokens);
    end
end


function issue = makeParsedIssue(typeName, sheetName, rowNumber, columnName, messageText)
    issue = struct( ...
        'Type', typeName, ...
        'Sheet', char(string(sheetName)), ...
        'Row', rowNumber, ...
        'Column', char(string(columnName)), ...
        'Message', char(string(messageText)));
end
