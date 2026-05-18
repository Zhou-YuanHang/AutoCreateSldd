function sldd_tool()
%% sldd_tool — AutoCreateSldd 图形界面 (App Designer 风格)
%  用法：sldd_tool

% 创建 Figure
fig = uifigure('Name', 'AutoCreateSldd v1.0', ...
    'Position', [500, 250, 580, 480], ...
    'Resize', 'off', ...
    'NumberTitle', 'off');

% 使用网格布局
gl = uigridlayout(fig, [7, 4], ...
    'RowHeight', {32, 32, 32, 28, 24, 24, '1x'}, ...
    'ColumnWidth', {85, '1x', 90, 90}, ...
    'Padding', [15, 12, 15, 12], ...
    'RowSpacing', 5, ...
    'ColumnSpacing', 8);

%% 第1行：方向
uilabel(gl, 'Text', '方向选择', ...
    'FontSize', 12, 'FontWeight', 'bold', ...
    'HorizontalAlignment', 'right', ...
    'Layout', [1, 1]);

dirGroup = uibuttongroup(gl, ...
    'BorderType', 'none', ...
    'BackgroundColor', [1 1 1], ...
    'Layout', [1, 2, 1, 3]);
radioE = uiradiobutton(dirGroup, 'Text', 'Excel → SLDD', ...
    'Position', [10, 5, 120, 22], 'Value', true);
radioS = uiradiobutton(dirGroup, 'Text', 'SLDD → Excel', ...
    'Position', [150, 5, 120, 22]);

%% 第2行：文件选择
uilabel(gl, 'Text', '数据文件', ...
    'FontSize', 12, 'FontWeight', 'bold', ...
    'HorizontalAlignment', 'right', ...
    'Layout', [2, 1]);

fld = uieditfield(gl, 'text', ...
    'Editable', 'off', ...
    'Placeholder', '自动选择当前文件夹第一个文件...', ...
    'Layout', [2, 2]);

uibutton(gl, 'Text', '浏览', ...
    'Layout', [2, 3], ...
    'ButtonPushedFcn', @(~,~) browseFile());

uibutton(gl, 'Text', '导出模板', ...
    'Layout', [2, 4], ...
    'ButtonPushedFcn', @(~,~) exportTemplate());

%% 第3行：操作
uibutton(gl, 'Text', '▶ 开始同步', ...
    'FontSize', 14, 'FontWeight', 'bold', ...
    'BackgroundColor', [0.25, 0.55, 0.95], ...
    'FontColor', [1 1 1], ...
    'Layout', [3, 1, 1, 4], ...
    'ButtonPushedFcn', @(~,~) runSync());

%% 第4行：文件信息
inf = uilabel(gl, 'Text', '', ...
    'FontSize', 11, ...
    'Layout', [4, 1, 1, 4]);

%% 第5行：校验提示
vld = uilabel(gl, 'Text', '', ...
    'FontSize', 11, ...
    'Layout', [5, 1, 1, 4]);

%% 第6行：状态
sta = uilabel(gl, 'Text', '就绪', ...
    'FontSize', 11, ...
    'FontColor', [0.4 0.4 0.4], ...
    'Layout', [6, 1, 1, 4]);

%% 第7行：日志
log = uitextarea(gl, ...
    'Editable', 'off', ...
    'FontName', 'Consolas', ...
    'FontSize', 10, ...
    'Placeholder', '操作日志...', ...
    'Layout', [7, 1, 1, 4]);

%% 数据
D.filePath = '';
D.valid = false;
init();

%% 回调
    function init()
        d = pwd;
        x = dir(fullfile(d, '*.xlsx'));
        if isempty(x), x = dir(fullfile(d, '*.xls')); end
        if ~isempty(x)
            D.filePath = fullfile(d, x(1).name);
            fld.Value = x(1).name;
            check();
        end
    end

    function browseFile()
        if radioE.Value
            [f, p] = uigetfile({'*.xlsx;*.xls', 'Excel 文件'});
        else
            [f, p] = uigetfile({'*.sldd', '数据字典'});
        end
        if f ~= 0
            D.filePath = fullfile(p, f);
            fld.Value = f;
            check();
        end
    end

    function exportTemplate()
        aLog('生成模板...');
        pyOk = false;
        try; [ok,~] = system('python generate_template.py'); pyOk = ok==0; catch; end
        if ~pyOk
            try; [ok,~] = system('python3 generate_template.py'); pyOk = ok==0; catch; end
        end
        if ~pyOk
            try; [ok,~] = system('"D:\Python\python.exe" generate_template.py'); pyOk = ok==0; catch; end
        end
        if pyOk
            aLog('模板已生成');
        else
            aLog('请手动运行 generate_template.py');
            uialert(fig, '请双击 generate_template.py 运行', '提示', 'Icon', 'info');
        end
    end

    function check()
        if isempty(D.filePath) || ~exist(D.filePath, 'file')
            D.valid = false; inf.Text = ''; vld.Text = ''; return;
        end
        d = dir(D.filePath);
        [~,n,e] = fileparts(D.filePath);
        inf.Text = sprintf('%s%s (%.1f KB)', n, e, d.bytes/1024);

        if ~radioE.Value
            D.valid = true;
            vld.Text = '✓ 就绪'; vld.FontColor = [0 0.6 0]; return;
        end

        try
            [~,sh] = xlsfinfo(D.filePath);
            err = {};
            for s = ["Signal","Parameter","Bus"]
                if ~any(strcmp(sh, s)), err{end+1}=sprintf('缺少 Sheet: %s', s); end
            end
            Chk = {
                'Signal',    ["VariableName","Package","Object","CustomStorageClass","DataType"];
                'Parameter', ["VariableName","Package","Object","CustomStorageClass","DataType","InitialValue"];
                'Bus',       ["BusName","Description","HeaderFile"];
                };
            for ci = 1:size(Chk,1)
                if ~any(strcmp(sh, Chk{ci,1})), continue; end
                try
                    o = detectImportOptions(D.filePath, 'Sheet', Chk{ci,1});
                    c = lower(o.VariableNames);
                    for ri = 1:length(Chk{ci,2})
                        if ~any(strcmp(c, lower(Chk{ci,2}{ri})))
                            err{end+1}=sprintf('%s 缺: %s', Chk{ci,1}, Chk{ci,2}{ri});
                        end
                    end
                catch ME
                    err{end+1}=sprintf('%s: %s', Chk{ci,1}, ME.message);
                end
            end
            if isempty(err)
                D.valid = true;
                vld.Text = '✓ 格式正确'; vld.FontColor = [0 0.6 0];
            else
                D.valid = false;
                vld.Text = '✗ 格式错误'; vld.FontColor = [0.8 0 0];
                for i = 1:length(err), aLog(err{i}); end
            end
        catch ME
            D.valid = false; vld.Text = ME.message; vld.FontColor = [0.8 0 0];
        end
    end

    function runSync()
        if isempty(D.filePath), aLog('请先选择文件'); return; end
        if ~D.valid && radioE.Value, aLog('文件校验未通过'); return; end
        sta.Text = '运行中...'; sta.FontColor = [0.6 0.6 0]; drawnow;
        try
            if radioE.Value, e2s(); else, s2e(); end
            sta.Text = '完成'; sta.FontColor = [0 0.6 0];
            check();
        catch ME
            sta.Text = '出错'; sta.FontColor = [0.8 0 0];
            aLog(ME.message);
        end
    end

    function e2s()
        [p,f] = fileparts(D.filePath);
        [~,sh] = xlsfinfo(D.filePath);
        oM = fullfile(p,[f,'_objects.m']);
        oS = fullfile(p,[f,'.sldd']);

        aLog('[1] 生成 M 脚本...');
        Excel2Workspace(sh, D.filePath, oM, [f,'.xlsx']);
        if ~exist(oM,'file'), aLog('生成失败'); return; end

        aLog('[2] 执行 M 脚本...');
        evalin('base', sprintf('run(''%s'')', oM));

        aLog('[3] 写入 SLDD...');
        v = evalin('base','who'); t = ["Simulink.Signal","Simulink.Parameter","Simulink.Bus"];
        n = {};
        for i = 1:length(v)
            try; x = evalin('base',v{i}); for j=1:length(t); if isa(x,t{j}); n{end+1}=v{i}; break; end; end; catch; end
        end
        if isempty(n), aLog('无 Simulink 对象'); return; end

        if exist(oS,'file'), d = Simulink.data.dictionary.open(oS);
        else, d = Simulink.data.dictionary.create(oS); end
        s = getSection(d,'Design Data');
        a=0;u=0;k=0;
        for i=1:length(n)
            nv = evalin('base',n{i});
            try
                e = getEntry(s,n{i});
                if isequal(struct(nv),struct(getValue(e))), k=k+1;
                else, setValue(e,nv); u=u+1; end
            catch, addEntry(s,n{i},nv); a=a+1; end
        end
        saveChanges(d); close(d);
        aLog(sprintf('+%d 改%d 跳%d',a,u,k));
    end

    function s2e()
        aLog('SLDD → Excel...');
        sldd_to_excel();
    end

    function aLog(m)
        c = log.Value;
        if isempty(c) || (iscell(c) && all(cellfun('isempty',c)))
            log.Value = {m};
        else
            log.Value = [c; {m}];
        end
        drawnow;
    end
end
