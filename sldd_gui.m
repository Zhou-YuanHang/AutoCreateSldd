function sldd_gui()
%% SLDD_GUI AutoCreateSldd 新一代图形界面
%  用法：sldd_gui
%
%  后端功能使用根目录的各函数（Excel2Workspace / validateDataDictionary / sldd_to_excel）
%  - Excel → M 脚本 → 执行 → 写入 .sldd
%  - SLDD → Excel 反向同步
%  - 校验 / 预览 / 模板生成

% ========== 主窗口 ==========
app = struct();
app.knownSheets = {'Signal', 'Parameter', 'Bus', 'BusElement'};

app.fig = figure( ...
    'Name', 'AutoCreateSldd v2.0 — 数据字典管理', ...
    'Tag', 'AutoCreateSlddGuiFigure', ...
    'NumberTitle', 'off', ...
    'MenuBar', 'none', ...
    'ToolBar', 'none', ...
    'Resize', 'off', ...
    'Position', [40, 40, 1380, 820], ...
    'Color', [0.94, 0.94, 0.94]);

% ========== Panel 1: 文件与输出 ==========
app.filePanel = uipanel(app.fig, ...
    'Title', '文件与输出', ...
    'FontWeight', 'bold', ...
    'Position', [0.015, 0.79, 0.97, 0.19]);

% --- 第1行：Excel 文件 ---
uicontrol(app.filePanel, 'Style', 'text', 'Position', [15, 92, 90, 22], ...
    'String', 'Excel 文件', 'HorizontalAlignment', 'left', 'FontWeight', 'bold');
app.excelPathEdit = uicontrol(app.filePanel, 'Style', 'edit', 'Tag', 'excelPathEdit', ...
    'Position', [105, 92, 980, 26], 'HorizontalAlignment', 'left', 'BackgroundColor', 'white');
app.browseExcelBtn = uicontrol(app.filePanel, 'Style', 'pushbutton', 'Position', [1100, 90, 90, 30], ...
    'String', '浏览...', 'Callback', @onBrowseExcel);
app.loadSheetsBtn = uicontrol(app.filePanel, 'Style', 'pushbutton', 'Tag', 'loadSheetsBtn', ...
    'Position', [1200, 90, 120, 30], 'String', '载入工作表', 'Callback', @onLoadSheets);

% --- 第2行：输出脚本 ---
uicontrol(app.filePanel, 'Style', 'text', 'Position', [15, 54, 90, 22], ...
    'String', '输出脚本', 'HorizontalAlignment', 'left', 'FontWeight', 'bold');
app.outputPathEdit = uicontrol(app.filePanel, 'Style', 'edit', 'Tag', 'outputPathEdit', ...
    'Position', [105, 54, 980, 26], 'HorizontalAlignment', 'left', 'BackgroundColor', 'white');
app.browseOutputBtn = uicontrol(app.filePanel, 'Style', 'pushbutton', 'Position', [1100, 52, 90, 30], ...
    'String', '另存为...', 'Callback', @onBrowseOutput);

% --- 第3行：数据字典 ---
uicontrol(app.filePanel, 'Style', 'text', 'Position', [15, 16, 90, 22], ...
    'String', '数据字典', 'HorizontalAlignment', 'left', 'FontWeight', 'bold');
app.slddPathEdit = uicontrol(app.filePanel, 'Style', 'edit', 'Tag', 'slddPathEdit', ...
    'Position', [105, 16, 980, 26], 'HorizontalAlignment', 'left', 'BackgroundColor', 'white');
app.browseSlddBtn = uicontrol(app.filePanel, 'Style', 'pushbutton', 'Position', [1100, 14, 90, 30], ...
    'String', '另存为...', 'Callback', @onBrowseSldd);
app.generateTemplateBtn = uicontrol(app.filePanel, 'Style', 'pushbutton', 'Position', [1200, 14, 120, 68], ...
    'String', '生成模板', 'Callback', @onGenerateTemplate, 'FontWeight', 'bold');

% ========== Panel 2: 工作表选择 ==========
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

% ========== Panel 3: 参数区 ==========
app.paramPanel = uipanel(app.fig, ...
    'Title', '参数区', ...
    'FontWeight', 'bold', ...
    'Position', [0.015, 0.21, 0.24, 0.20]);

app.validateBeforeRunChk = uicontrol(app.paramPanel, 'Style', 'checkbox', 'Position', [12, 108, 210, 24], ...
    'String', '执行前先做数据校验', 'Value', 1, 'BackgroundColor', [0.94, 0.94, 0.94]);
app.confirmBeforeSaveChk = uicontrol(app.paramPanel, 'Style', 'checkbox', 'Position', [12, 82, 210, 24], ...
    'String', '写入字典前弹确认框', 'Value', 1, 'BackgroundColor', [0.94, 0.94, 0.94]);
app.onlySaveGeneratedChk = uicontrol(app.paramPanel, 'Style', 'checkbox', 'Position', [12, 56, 280, 24], ...
    'String', '同步时删除SLDD中多余的变量（危险）', 'Value', 0, 'BackgroundColor', [0.94, 0.94, 0.94]);

uicontrol(app.paramPanel, 'Style', 'text', 'Position', [12, 22, 120, 22], ...
    'String', '工作表预览行数', 'HorizontalAlignment', 'left');
app.previewRowsEdit = uicontrol(app.paramPanel, 'Style', 'edit', 'Position', [110, 22, 60, 24], ...
    'String', '12', 'BackgroundColor', 'white', 'HorizontalAlignment', 'center', ...
    'Callback', @onPreviewRowCountChanged);
uicontrol(app.paramPanel, 'Style', 'text', 'Position', [185, 22, 118, 22], ...
    'String', '错误定位会围绕该行预览', 'HorizontalAlignment', 'left', 'ForegroundColor', [0.4, 0.4, 0.4]);

% ========== Panel 4: 操作 ==========
app.actionPanel = uipanel(app.fig, ...
    'Title', '操作', ...
    'FontWeight', 'bold', ...
    'Position', [0.015, 0.04, 0.24, 0.15]);

% 行1: 校验 + 预览
app.validateBtn = uicontrol(app.actionPanel, 'Style', 'pushbutton', 'Tag', 'validateBtn', ...
    'Position', [12, 78, 145, 32], 'String', '校验所选工作表', 'Callback', @onValidate, 'FontWeight', 'bold');
app.previewSheetBtn = uicontrol(app.actionPanel, 'Style', 'pushbutton', 'Position', [167, 78, 145, 32], ...
    'String', '预览所选工作表', 'Callback', @onPreviewCurrentSheet, 'FontWeight', 'bold');

% 行2: 仅生成 + 完整导入
app.generateBtn = uicontrol(app.actionPanel, 'Style', 'pushbutton', 'Tag', 'generateBtn', ...
    'Position', [12, 40, 145, 32], 'String', '仅生成 objects', 'Callback', @onGenerateOnly, 'FontWeight', 'bold');
app.importBtn = uicontrol(app.actionPanel, 'Style', 'pushbutton', 'Tag', 'importBtn', ...
    'Position', [167, 40, 145, 32], 'String', '完整导入到 SLDD', 'Callback', @onFullImport, ...
    'FontWeight', 'bold', 'ForegroundColor', [0.1, 0.1, 0.1], 'BackgroundColor', [0.82, 0.9, 0.82]);

% 行3: 反向同步（满宽）
app.reverseSyncBtn = uicontrol(app.actionPanel, 'Style', 'pushbutton', 'Tag', 'reverseSyncBtn', ...
    'Position', [12, 4, 300, 30], 'String', '↔ SLDD → Excel 反向同步', ...
    'Callback', @onReverseSync, 'FontWeight', 'bold');

% ========== Tab Group: 预览/结果/日志 ==========
app.tabGroup = uitabgroup(app.fig, 'Position', [0.27, 0.24, 0.715, 0.51]);

% Tab 1: 工作表预览
app.sheetPreviewTab = uitab(app.tabGroup, 'Title', '工作表预览');
uicontrol(app.sheetPreviewTab, 'Style', 'text', 'Position', [12, 326, 80, 22], ...
    'String', '预览工作表', 'HorizontalAlignment', 'left', 'FontWeight', 'bold');
app.previewSheetPopup = uicontrol(app.sheetPreviewTab, 'Style', 'popupmenu', 'Position', [92, 326, 220, 24], ...
    'BackgroundColor', 'white', 'String', {'(未载入)'}, 'Callback', @onPreviewCurrentSheet);
app.previewSummaryText = uicontrol(app.sheetPreviewTab, 'Style', 'text', 'Position', [330, 320, 590, 32], ...
    'String', {'尚未预览工作表'}, 'HorizontalAlignment', 'left', 'BackgroundColor', [0.94, 0.94, 0.94]);
app.sheetPreviewTable = uitable(app.sheetPreviewTab, 'Position', [12, 12, 928, 302], ...
    'Data', cell(0, 1), 'ColumnName', {'预览结果'}, 'RowName', []);

% Tab 2: 结果预览
app.resultPreviewTab = uitab(app.tabGroup, 'Title', '结果预览');
app.resultSummaryText = uicontrol(app.resultPreviewTab, 'Style', 'text', 'Position', [12, 312, 928, 40], ...
    'String', {'尚未生成结果'}, 'HorizontalAlignment', 'left', 'BackgroundColor', [0.94, 0.94, 0.94]);
app.resultTable = uitable(app.resultPreviewTab, 'Position', [12, 12, 928, 292], ...
    'Data', cell(0, 4), 'ColumnName', {'变量名', '对象类型', '构造器', '状态'}, 'RowName', []);



% ========== 运行日志 Panel ==========
app.logPanel = uipanel(app.fig, ...
    'Title', '运行日志', ...
    'FontWeight', 'bold', ...
    'Position', [0.27, 0.04, 0.715, 0.20]);
app.logBox = uicontrol(app.logPanel, 'Style', 'edit', 'Position', [12, 12, 840, 130], ...
    'Max', 200, 'Min', 0, 'HorizontalAlignment', 'left', 'BackgroundColor', 'white');
app.clearLogBtn = uicontrol(app.logPanel, 'Style', 'pushbutton', 'Position', [865, 102, 75, 28], ...
    'String', '清空日志', 'Callback', @onClearLog);

% ========== 状态栏 ==========
app.statusText = uicontrol(app.fig, 'Style', 'text', 'Position', [20, 8, 1320, 22], ...
    'String', '就绪', 'HorizontalAlignment', 'left', 'BackgroundColor', [0.94, 0.94, 0.94], ...
    'ForegroundColor', [0.2, 0.2, 0.2], 'FontWeight', 'bold');

% ========== 应用数据初始 ==========
setappdata(app.fig, 'AllSheets', {});
setappdata(app.fig, 'LastGeneratedInfo', struct('Name', {}, 'Constructor', {}, 'Type', {}));

appendLog('AutoCreateSldd v2.0 GUI 已启动。');
setStatus('就绪');

tryLoadDefaultWorkbook();

% ==================================================================
%  回调 — 文件操作
% ==================================================================

    function onBrowseExcel(~, ~)
        [fileName, filePath] = uigetfile({'*.xlsx;*.xls', 'Excel 文件 (*.xlsx, *.xls)'}, '选择数据字典 Excel');
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
            clearPreviewAndLog();
            loadSheetsFromExcel(excelPath);
            appendLog(['已载入 Excel：', excelPath]);
            setStatus(sprintf('已载入：%s', getFileName(excelPath)));
        catch ME
            appendLog(['载入失败：', ME.message]);
            for si = 1:min(length(ME.stack), 8)
                appendLog(['  → ', ME.stack(si).name, ' (行', num2str(ME.stack(si).line), ')']);
            end
            setStatus('载入 Excel 失败', true);
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
        [fileName, filePath] = uiputfile({'*.m', 'MATLAB 脚本 (*.m)'}, '输出脚本位置', basePath);
        if isequal(fileName, 0); return; end
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
        [fileName, filePath] = uiputfile({'*.sldd', 'Simulink 数据字典 (*.sldd)'}, '数据字典位置', basePath);
        if isequal(fileName, 0); return; end
        set(app.slddPathEdit, 'String', fullfile(filePath, fileName));
        appendLog('已更新数据字典路径。');
    end

    function onGenerateTemplate(~, ~)
        excelPath = strtrim(get(app.excelPathEdit, 'String'));
        if isempty(excelPath)
            defaultTarget = fullfile(pwd, 'MCU_Template.xlsx');
        else
            [folder, ~] = fileparts(excelPath);
            defaultTarget = fullfile(folder, 'MCU_Template.xlsx');
        end

        [fileName, filePath] = uiputfile({'*.xlsx', 'Excel 文件 (*.xlsx)'}, '生成模板到', defaultTarget);
        if isequal(fileName, 0); appendLog('已取消生成模板。'); return; end
        fullOutPath = fullfile(filePath, fileName);

        % Python 版模板生成（传入输出路径，跳过 Python 的弹窗）
        pyOk = false;
        try; [ok, ~] = system(sprintf('python generate_template.py "%s"', fullOutPath)); pyOk = (ok == 0); catch; end
        if ~pyOk
            try; [ok, ~] = system(sprintf('python3 generate_template.py "%s"', fullOutPath)); pyOk = (ok == 0); catch; end
        end
        if pyOk
            appendLog(['模板已生成：', fullOutPath]);
            set(app.excelPathEdit, 'String', fullOutPath);
            loadWorkbook(fullOutPath);
        else
            appendLog('无法生成模板，请手动运行 generate_template.py。');
        end
    end

% ==================================================================
%  回调 — 工作表选择
% ==================================================================

    function onSelectRecommended(~, ~)
        allSheets = getAllSheets();
        if isempty(allSheets); appendLog('尚未载入工作表。'); return; end
        selectIdx = find(ismember(allSheets, app.knownSheets));
        if isempty(selectIdx); selectIdx = 1:numel(allSheets); end
        set(app.sheetList, 'Value', selectIdx);
        onSheetSelectionChanged();
        appendLog('已按标准工作表预选。');
    end

    function onSelectAll(~, ~)
        allSheets = getAllSheets();
        if isempty(allSheets); return; end
        set(app.sheetList, 'Value', 1:numel(allSheets));
        onSheetSelectionChanged();
        appendLog('已全选所有工作表。');
    end

    function onClearSelection(~, ~)
        set(app.sheetList, 'Value', []);
        updatePreviewSheetChoices({});
        resetSheetPreview();
        setStatus('已清空工作表选择。');
        appendLog('已清空工作表选择。');
    end

    function onSheetSelectionChanged(~, ~)
        selectedSheets = safeGetSelectedSheets();
        updatePreviewSheetChoices(selectedSheets);
        if isempty(selectedSheets)
            resetSheetPreview();
            return;
        end
    end

% ==================================================================
%  回调 — 操作
% ==================================================================

    function onValidate(~, ~)
        try
            excelPath = getRequiredFilePath(app.excelPathEdit, '请先选择 Excel 文件。');
            selectedSheets = getSelectedSheets();
            [isValid, issues] = validateDataDictionary(excelPath, selectedSheets);

            appendLog(['校验工作表：', strjoin(selectedSheets, ', ')]);
            if isValid
                appendLog('✅ 校验通过，未发现问题。');
                setStatus('校验通过。');
                appendLog('校验通过 ✅ 未发现结构性问题。');
            else
                appendLog(sprintf('❌ 校验失败，共 %d 条问题：', numel(issues)));
                for ii = 1:numel(issues)
                    appendLog(['  ', issues{ii}]);
                end
                setStatus('校验失败。', true);
                appendLog(sprintf('校验失败 ❌ %d 条问题', numel(issues)));
            end
        catch ME
            appendLog(['校验异常：', ME.message]);
            setStatus('校验异常。', true);
        end
    end

    function onPreviewCurrentSheet(~, ~)
        previewSheetName = getCurrentPreviewSheetName();
        if isempty(previewSheetName)
            resetSheetPreview();
            return;
        end
        showSheetPreview(previewSheetName);
    end

    function onGenerateOnly(~, ~)
        try
            [outputPath, selectedSheets, generatedInfo] = doGenerateObjects();
            appendLog(['objects 已生成：', outputPath]);
            setStatus('生成完成。');
            showGeneratedResult(generatedInfo, selectedSheets, false);
        catch ME
            appendLog(['生成失败：', ME.message]);
            setStatus('生成失败。', true);
        end
    end

    function onReverseSync(~, ~)
        % SLDD → Excel 反向同步，直接调用 sldd_to_excel()
        try
            appendLog('启动 SLDD → Excel 反向同步...');
            % sldd_to_excel 在根目录，自动在 path 上
            sldd_to_excel();
            appendLog('反向同步完成。');
            setStatus('反向同步完成 ✅');
        catch ME
            appendLog(['反向同步失败：', ME.message]);
            setStatus('反向同步失败。', true);
        end
    end

    function onFullImport(~, ~)
        try
            % 1) 生成
            [outputPath, selectedSheets, generatedInfo] = doGenerateObjects();
            appendLog('[1/3] 生成 M 脚本完成。');

            % 2) 执行
            scriptContent = fileread(outputPath);
            evalin('base', scriptContent);
            appendLog('[2/3] 脚本已执行到基础工作区。');

            % 3) 写入 SLDD
            slddPath = getRequiredFilePath(app.slddPathEdit, '请先填写数据字典路径。');
            confirmSave = logical(get(app.confirmBeforeSaveChk, 'Value'));
            onlySaveGenerated = logical(get(app.onlySaveGeneratedChk, 'Value'));

            if onlySaveGenerated
                varNames = {generatedInfo.Name};
                appendLog(sprintf('仅写入本次生成的 %d 个变量。', numel(varNames)));
            else
                varNames = {};
                appendLog('写入基础工作区全部 Simulink 对象。');
            end

            doWriteToSldd(slddPath, varNames, confirmSave);
            appendLog('[3/3] 数据字典写入完成。');
            showGeneratedResult(generatedInfo, selectedSheets, true);
            setStatus('导入完成 ✅');
        catch ME
            appendLog(['导入失败：', ME.message]);
            setStatus('导入失败。', true);
        end
    end

% ==================================================================
%  回调 — 日志
% ==================================================================

    function onClearLog(~, ~)
        set(app.logBox, 'String', {});
        appendLog('日志已清空。');
    end

    function onClearIssues(~, ~)
        appendLog('消息区已清空。');
    end

% ==================================================================
%  核心流程函数
% ==================================================================

    function [outputPath, selectedSheets, generatedInfo] = doGenerateObjects()
        excelPath = getRequiredFilePath(app.excelPathEdit, '请先选择 Excel 文件。');
        outputPath = getRequiredFilePath(app.outputPathEdit, '请先填写输出脚本路径。');
        selectedSheets = getSelectedSheets();
        allSheets = getAllSheets();

        % 可选校验
        if logical(get(app.validateBeforeRunChk, 'Value'))
            appendLog('执行前校验...');
            [isValid, issues] = validateDataDictionary(excelPath, selectedSheets);
            if ~isValid
                appendLog('❌ 校验未通过，已停止：');
                for ii = 1:numel(issues); appendLog(['  ', issues{ii}]); end
                error('数据校验未通过，请修正后重试。');
            end
            appendLog('✅ 校验通过。');
        end

        % 确保输出路径有 .m 后缀
        [folder, ~, ext] = fileparts(outputPath);
        if isempty(ext)
            outputPath = [outputPath, '.m'];
            set(app.outputPathEdit, 'String', outputPath);
            [folder, ~] = fileparts(outputPath);
        end
        if ~isempty(folder) && ~exist(folder, 'dir'); mkdir(folder); end

        % 生成 M 脚本（通过 Excel2Workspace，传入 sheet 索引跳过 listdlg）
        appendLog(sprintf('生成 M 脚本：%s', outputPath));
        appendLog(sprintf('工作表：%s', strjoin(selectedSheets, ', ')));
        [~, xlsxBase] = fileparts(excelPath);
        selectedIndices = find(ismember(allSheets, selectedSheets));
        Excel2Workspace(allSheets, excelPath, outputPath, [xlsxBase, '.xlsx'], selectedIndices);

        % 收集生成的对象信息
        generatedInfo = collectGeneratedObjectInfo(outputPath);
        generatedInfo = filterExportedObjectInfo(generatedInfo);
        setappdata(app.fig, 'LastGeneratedInfo', generatedInfo);
    end

    function doWriteToSldd(slddPath, targetVarNames, confirmSave)
        % 确保 .sldd 后缀
        [dictPath, dictName, ext] = fileparts(slddPath);
        if isempty(ext); slddPath = [slddPath, '.sldd']; end
        if isempty(dictPath); dictPath = pwd; slddPath = fullfile(dictPath, [dictName, '.sldd']); end

        % 获取基础工作区变量
        allVars = evalin('base', 'who');
        excludeVars = {'ans'};
        allVars = setdiff(allVars, excludeVars, 'stable');

        if isempty(targetVarNames)
            % 筛选所有 Simulink 对象变量
            varNames = {};
            for ii = 1:numel(allVars)
                try
                    v = evalin('base', allVars{ii});
                    if isa(v, 'Simulink.Signal') || isa(v, 'Simulink.Parameter') || isa(v, 'Simulink.Bus')
                        varNames{end+1} = allVars{ii}; %#ok<AGROW>
                    end
                catch
                end
            end
            appendLog(sprintf('基础工作区中找到 %d 个 Simulink 对象。', numel(varNames)));
        else
            varNames = targetVarNames(ismember(targetVarNames, allVars));
            missing = targetVarNames(~ismember(targetVarNames, allVars));
            if ~isempty(missing)
                appendLog(sprintf('未找到（已跳过）：%s', strjoin(missing, ', ')));
            end
        end

        if isempty(varNames)
            error('没有可写入的变量。请先生成并执行脚本。');
        end

        % 确认对话框
        if confirmSave
            previewCount = min(numel(varNames), 8);
            previewText = strjoin(varNames(1:previewCount), newline);
            if numel(varNames) > previewCount
                previewText = sprintf('%s\n... 共 %d 个', previewText, numel(varNames));
            end
            choice = questdlg(sprintf('即将写入 %d 个变量到数据字典：\n\n%s', numel(varNames), previewText), ...
                '确认写入', '继续', '取消', '继续');
            if ~strcmp(choice, '继续')
                appendLog('用户取消写入。');
                return;
            end
        end

        % 打开/创建 SLDD
        if exist(slddPath, 'file')
            dictObj = Simulink.data.dictionary.open(slddPath);
            appendLog(['已打开：', getFileName(slddPath)]);
        else
            dictObj = Simulink.data.dictionary.create(slddPath);
            appendLog(['已创建：', getFileName(slddPath)]);
        end
        closeGuard = onCleanup(@() safeCloseDict(dictObj));

        dataSect = getSection(dictObj, 'Design Data');

        % 如果仅写入生成变量，先清理不在目标列表中的旧条目
        if ~isempty(targetVarNames)
            [~, slddName] = fileparts(slddPath);
            removed = cleanManagedEntries(dataSect, targetVarNames, [slddName, '.sldd']);
            if removed > 0; appendLog(sprintf('已清理 %d 条旧条目。', removed)); end
        end

        % 逐条写入（通过 DataSource 过滤，只操作当前字典）
        [~, slddFileName] = fileparts(slddPath);
        added = 0; updated = 0; skipped = 0; failed = 0;
        for ii = 1:numel(varNames)
            vn = varNames{ii};
            try
                newVal = evalin('base', vn);
                % 在当前字典中查找同名条目
                allE = find(dataSect);
                localMask = strcmp({allE.DataSource}, [slddFileName, '.sldd']) & strcmp({allE.Name}, vn);
                if any(localMask)
                    entryObj = allE(localMask);
                    oldVal = getValue(entryObj);
                    if isEqualByStruct(newVal, oldVal)
                        skipped = skipped + 1;
                    else
                        setValue(entryObj, newVal);
                        updated = updated + 1;
                    end
                else
                    addEntry(dataSect, vn, newVal);
                    added = added + 1;
                end
            catch ME
                failed = failed + 1;
                appendLog(['  ✗ ', vn, '：', ME.message]);
            end
        end

        saveChanges(dictObj);
        appendLog(sprintf('写入完成：+%d ～%d -%d ✗%d', added, updated, skipped, failed));

        % 验证
        try
            vDict = Simulink.data.dictionary.open(slddPath);
            vSect = getSection(vDict, 'Design Data');
            entries = find(vSect);
            appendLog(sprintf('验证：SLDD 共 %d 个条目（含引用）。', numel(entries)));
            safeCloseDict(vDict);
        catch ME
            appendLog(['验证失败：', ME.message]);
        end
    end

% ==================================================================
%  辅助：UI 更新
% ==================================================================

    function loadSheetsFromExcel(excelPath)
        if ~exist(excelPath, 'file'); error('文件不存在：%s', excelPath); end
        allSheets = cellstr(sheetnames(excelPath));
        setappdata(app.fig, 'AllSheets', allSheets);
        defaultSel = defaultSelection(allSheets);
        set(app.sheetList, 'String', allSheets, 'Value', defaultSel);
        updatePreviewSheetChoices(allSheets(defaultSel));
        setStatus(sprintf('已载入 %d 个工作表', numel(allSheets)));
        if ~isempty(allSheets)
            showSheetPreview(allSheets{defaultSel(1)});
        end
    end

    function tryLoadDefaultWorkbook()
        xlsxFiles = dir(fullfile(pwd, '*.xlsx'));
        if isempty(xlsxFiles); xlsxFiles = dir(fullfile(pwd, '*.xls')); end
        if isempty(xlsxFiles); return; end
        candidate = fullfile(pwd, xlsxFiles(1).name);
        set(app.excelPathEdit, 'String', candidate);
        try
            loadWorkbook(candidate);
        catch ME
            appendLog(['自动载入失败：', ME.message]);
            for si = 1:min(length(ME.stack), 5)
                appendLog(['  → ', ME.stack(si).name, ' (行', num2str(ME.stack(si).line), ')']);
            end
        end
    end

    function allSheets = getAllSheets()
        allSheets = getappdata(app.fig, 'AllSheets');
        if isempty(allSheets); allSheets = {}; end
    end

    function selectedSheets = getSelectedSheets()
        allSheets = getAllSheets();
        if isempty(allSheets); error('请先载入工作表。'); end
        selIdx = get(app.sheetList, 'Value');
        if isempty(selIdx); error('请至少选择一个工作表。'); end
        selectedSheets = allSheets(selIdx);
    end

    function selectedSheets = safeGetSelectedSheets()
        try; selectedSheets = getSelectedSheets(); catch; selectedSheets = {}; end
    end

    function updatePreviewSheetChoices(candidateSheets)
        allSheets = getAllSheets();
        if nargin < 1 || isempty(candidateSheets); candidateSheets = allSheets; end
        if isempty(candidateSheets)
            set(app.previewSheetPopup, 'String', {'(未载入)'}, 'Value', 1);
            return;
        end
        currentName = getCurrentPreviewSheetName();
        set(app.previewSheetPopup, 'String', candidateSheets);
        idx = find(strcmp(candidateSheets, currentName), 1);
        if isempty(idx); idx = 1; end
        set(app.previewSheetPopup, 'Value', idx);
    end

    function sheetName = getCurrentPreviewSheetName()
        items = get(app.previewSheetPopup, 'String');
        if ischar(items); items = {items}; end
        v = get(app.previewSheetPopup, 'Value');
        v = min(max(1, v), numel(items));
        sheetName = items{v};
        if strcmp(sheetName, '(未载入)'); sheetName = ''; end
    end

    function showSheetPreview(sheetName, focusRow)
        excelPath = strtrim(get(app.excelPathEdit, 'String'));
        if isempty(excelPath) || isempty(sheetName); resetSheetPreview(); return; end
        previewRows = getPreviewRowCount();
        [dataTable, readIssue] = readSheetAsStringsSimple(excelPath, sheetName);
        if ~isempty(readIssue)
            set(app.previewSummaryText, 'String', {['读取失败：', readIssue]});
            setStatus('预览失败。', true);
            return;
        end
        % focusRow 可能未传入（nargin < 2），不能直接引用
        if nargin >= 2
            [previewData, previewColumns, summaryLines] = tableToPreview(dataTable, previewRows, focusRow);
        else
            [previewData, previewColumns, summaryLines] = tableToPreview(dataTable, previewRows);
        end
        set(app.sheetPreviewTable, 'Data', previewData, 'ColumnName', previewColumns);
        set(app.previewSummaryText, 'String', summaryLines);
        set(app.tabGroup, 'SelectedTab', app.sheetPreviewTab);
    end

    function resetSheetPreview()
        set(app.sheetPreviewTable, 'Data', cell(0, 1), 'ColumnName', {'预览结果'});
        set(app.previewSummaryText, 'String', {'尚未预览工作表'});
    end

    function showGeneratedResult(generatedInfo, selectedSheets, includeExecStatus)
        if isempty(generatedInfo)
            set(app.resultTable, 'Data', cell(0, 4), 'ColumnName', {'变量名','对象类型','构造器','状态'});
            set(app.resultSummaryText, 'String', {'本次没有识别到可导入对象。'});
            return;
        end
        resultData = cell(numel(generatedInfo), 4);
        for ii = 1:numel(generatedInfo)
            resultData{ii, 1} = generatedInfo(ii).Name;
            resultData{ii, 2} = generatedInfo(ii).Type;
            resultData{ii, 3} = generatedInfo(ii).Constructor;
            if includeExecStatus
                try
                    allV = evalin('base', 'who');
                    if ismember(generatedInfo(ii).Name, allV)
                        vv = evalin('base', generatedInfo(ii).Name);
                        resultData{ii, 4} = class(vv);
                    else
                        resultData{ii, 4} = '未找到';
                    end
                catch
                    resultData{ii, 4} = '未知';
                end
            else
                resultData{ii, 4} = '待执行';
            end
        end
        set(app.resultTable, 'Data', resultData, ...
            'ColumnName', {'变量名','对象类型','构造器','状态'});
        types = {generatedInfo.Type};
        uTypes = unique(types, 'stable');
        typeSum = cell(1, numel(uTypes));
        for ii = 1:numel(uTypes)
            typeSum{ii} = sprintf('%s=%d', uTypes{ii}, sum(strcmp(types, uTypes{ii})));
        end
        summary = sprintf('工作表：%s | 对象数：%d | %s', ...
            strjoin(selectedSheets, ', '), numel(generatedInfo), strjoin(typeSum, '  '));
        set(app.resultSummaryText, 'String', {summary, ...
            '状态列显示当前基础工作区中的实际类（已执行）或"待执行"（仅生成）。'});
        set(app.tabGroup, 'SelectedTab', app.resultPreviewTab);
    end



    function appendLog(msg)
        c = get(app.logBox, 'String');
        if isempty(c) || (iscell(c) && all(cellfun('isempty', c)))
            set(app.logBox, 'String', {msg});
        else
            set(app.logBox, 'String', [c; {msg}]);
        end
        drawnow;
    end

    function setStatus(msg, isError)
        if nargin < 2; isError = false; end
        if isError
            set(app.statusText, 'String', ['⚠ ', msg], 'ForegroundColor', [0.8, 0, 0]);
        else
            set(app.statusText, 'String', msg, 'ForegroundColor', [0.2, 0.2, 0.2]);
        end
        drawnow;
    end

% ==================================================================
%  实用工具函数
% ==================================================================

    function refreshDerivedPaths(excelPath)
        [folder, baseName] = fileparts(excelPath);
        if isempty(folder); folder = pwd; end
        % 自动填充输出脚本路径
        outPath = get(app.outputPathEdit, 'String');
        if isempty(outPath)
            set(app.outputPathEdit, 'String', fullfile(folder, [baseName, '_objects.m']));
        end
        % 自动填充数据字典路径
        slddPath = get(app.slddPathEdit, 'String');
        if isempty(slddPath)
            set(app.slddPathEdit, 'String', fullfile(folder, [baseName, '.sldd']));
        end
    end

    function clearPreviewAndLog()
        resetSheetPreview();
        set(app.resultTable, 'Data', cell(0, 4));
        set(app.resultSummaryText, 'String', {'尚未生成结果'});
    end

    function fp = getRequiredFilePath(editHandle, errMsg)
        fp = strtrim(get(editHandle, 'String'));
        if isempty(fp); error(errMsg); end
    end

    function n = getPreviewRowCount()
        s = get(app.previewRowsEdit, 'String');
        % MATLAB R2024b 可能返回 string 而非 char
        if isstring(s); s = char(s); end
        s = strtrim(s);
        n = str2double(s);
        if isnan(n) || n < 1; n = 12; end
    end

    function onPreviewRowCountChanged(~, ~)
        % 修改预览行数后自动刷新当前预览
        previewSheetName = getCurrentPreviewSheetName();
        if isempty(previewSheetName); return; end
        showSheetPreview(previewSheetName);
    end

    function sel = defaultSelection(allSheets, ensureMin)
        if nargin < 2; ensureMin = 0; end
        sel = find(ismember(allSheets, app.knownSheets));
        if isempty(sel) && ensureMin > 0; sel = 1; end
    end

    function name = getFileName(fullpath)
        [~, n, e] = fileparts(fullpath);
        name = [n, e];
    end

    function [dataTable, readIssue] = readSheetAsStringsSimple(excelPath, sheetName)
        try
            opts = detectImportOptions(excelPath, 'Sheet', sheetName);
            opts.VariableNamingRule = 'preserve';
            opts = setvartype(opts, opts.VariableNames, 'string');
            dataTable = readtable(excelPath, opts);
            readIssue = '';
        catch ME
            dataTable = table();
            readIssue = ME.message;
        end
    end

    function [previewData, previewColumns, summaryLines] = tableToPreview(sheetTable, previewRows, focusRow)
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
            if isempty(nearestIdx); nearestIdx = totalRows; end
            halfWindow = floor(previewRows / 2);
            startIdx = max(1, nearestIdx - halfWindow);
            endIdx = min(totalRows, startIdx + previewRows - 1);
            startIdx = max(1, endIdx - previewRows + 1);
            focusText = sprintf('已围绕 Excel 第 %d 行定位，显示第 %d-%d 行。', focusRow, excelRowNumbers(startIdx), excelRowNumbers(endIdx));
        end
        previewSlice = sheetTable(startIdx:endIdx, :);
        previewCells = table2cell(previewSlice);
        previewCells = sanitizeCells(previewCells);
        previewData = cell(size(previewCells, 1), size(previewCells, 2) + 1);
        previewData(:, 1) = num2cell(excelRowNumbers(startIdx:endIdx));
        previewData(:, 2:end) = previewCells;
        summaryLines = {sprintf('总数据行：%d，当前预览：%d 行。', totalRows, size(previewData, 1)), focusText};
    end

    function cellData = sanitizeCells(cellData)
        for rr = 1:size(cellData, 1)
            for cc = 1:size(cellData, 2)
                v = cellData{rr, cc};
                if isstring(v)
                    if ismissing(v); cellData{rr, cc} = ''; else; cellData{rr, cc} = char(v); end
                elseif isempty(v)
                    cellData{rr, cc} = '';
                elseif ~ischar(v) && ~isnumeric(v) && ~islogical(v)
                    cellData{rr, cc} = char(string(v));
                end
            end
        end
    end

    function info = collectGeneratedObjectInfo(mFilePath)
        scriptContent = fileread(mFilePath);
        lines = regexp(scriptContent, '\r\n|\n|\r', 'split');
        info = struct('Name', {}, 'Constructor', {}, 'Type', {});
        seen = {};
        for ii = 1:numel(lines)
            tokens = regexp(lines{ii}, '^\s*([A-Za-z]\w*)\s*=\s*([A-Za-z]\w*(?:\.[A-Za-z]\w*)*)\s*;\s*$', 'tokens', 'once');
            if isempty(tokens); continue; end
            vn = tokens{1}; ct = tokens{2};
            if any(strcmp(seen, vn)); continue; end
            seen{end+1} = vn;
            info(end+1).Name = vn; %#ok<AGROW>
            info(end).Constructor = ct;
            dotIdx = find(ct == '.', 1, 'last');
            if isempty(dotIdx); info(end).Type = ct; else; info(end).Type = ct(dotIdx+1:end); end
        end
    end

    function info = filterExportedObjectInfo(info)
        keep = true(1, numel(info));
        for ii = 1:numel(info)
            name = info(ii).Name;
            keep(ii) = ~strcmp(info(ii).Type, 'BusElement') ...
                && ~(numel(name) >= 9 && strcmp(name(end-8:end), '_elements'));
        end
        info = info(keep);
    end

    function eq = isEqualByStruct(a, b)
        if ~strcmp(class(a), class(b)); eq = false; return; end
        try; eq = isequal(struct(a), struct(b)); catch; try; eq = isequal(a, b); catch; eq = false; end; end
    end

    function n = cleanManagedEntries(dataSect, keepNames, slddName)
        allE = find(dataSect);
        localMask = strcmp({allE.DataSource}, slddName);
        entries = allE(localMask);
        n = 0;
        for ii = 1:numel(entries)
            en = entries(ii).Name;
            if ismember(en, keepNames); continue; end
            try
                ev = getValue(entries(ii));
                if isa(ev, 'Simulink.Signal') || isa(ev, 'Simulink.Parameter') || isa(ev, 'Simulink.Bus')
                    deleteEntry(dataSect, en);
                    n = n + 1;
                end
            catch
            end
        end
    end

    function safeCloseDict(dictObj)
        try; close(dictObj); catch; end
    end

end
