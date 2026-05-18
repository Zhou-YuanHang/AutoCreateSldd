function templatePath = generateTemplate(outputDir, templateName)
% GENERATETEMPLATE 生成标准数据字典 Excel 模板
%
%   templatePath = generateTemplate()
%   templatePath = generateTemplate(outputDir)
%   templatePath = generateTemplate(outputDir, templateName)
%
% 输入参数：
%   outputDir    - 模板输出目录，默认当前工作目录
%   templateName - 模板文件名，默认 MCU_Template.xlsx
%
% 输出参数：
%   templatePath - 实际生成的模板完整路径
%
% 说明：
%   本函数通过复制内置种子文件（AutoCreateSldd_TemplateSeed.xlsx）生成模板。
%   种子文件由 Python openpyxl 预构建，已包含：
%     - 6 个工作表（History、Signal、Parameter、Config、Bus、BusElement）
%     - 示例数据
%     - 表头颜色（深蓝必填、浅蓝选填、中蓝 Config）
%     - 示例行浅绿底色
%     - 列宽优化、冻结首行
%     - 16 处数据验证下拉列表（引用 Config 表）
%     - 4 处批注
%   复制后会自动更新 History 工作表中的日期。

if nargin < 1 || isempty(outputDir)
    outputDir = pwd;
end
if nargin < 2 || isempty(templateName)
    templateName = 'MCU_Template.xlsx';
end

if ~exist(outputDir, 'dir')
    mkdir(outputDir);
end

[~, nameOnly, ext] = fileparts(templateName);
if isempty(ext)
    ext = '.xlsx';
end

templatePath = getUniqueFilePath(fullfile(outputDir, [nameOnly, ext]));

% ---- 定位种子文件 ----
seedPath = getSeedPath();
if isempty(seedPath) || ~exist(seedPath, 'file')
    error('AutoCreateSldd:seedNotFound', ...
        '种子模板文件不存在：%s\n请确保 AutoCreateSldd_TemplateSeed.xlsx 与 generateTemplate.m 在同一目录下。', ...
        getExpectedSeedPath());
end

% ---- 复制种子文件 ----
fprintf('正在生成模板...\n');
copyfile(seedPath, templatePath);

% ---- 更新 History 日期 ----
try
    todayStr = char(string(datetime('today', 'Format', 'yyyy-MM-dd')));
    historyData = {'Version', 'ChangeDate', 'Changer', 'Content'; ...
                   'V1.0', todayStr, 'AutoCreateSldd', '模板初始化'; ...
                   '说明', '蓝色表头为必填列，浅蓝色表头为选填列', '', ...
                   '模板已内置示例数据与 Config 下拉列表，可直接改名改值后使用'};
    writecell(historyData, templatePath, 'Sheet', 'History');
catch ME
    warning('更新 History 日期失败：%s', ME.message);
end

fprintf('\n===============================================\n');
fprintf('模板生成完成\n');
fprintf('===============================================\n');
fprintf('模板路径: %s\n', templatePath);
fprintf('模板已包含示例数据、格式和 Config 下拉配置。\n');
fprintf('请先填写 Signal / Parameter / Bus / BusElement 工作表，再运行 main() 导入。\n');

end


% ==================================================================
%  种子文件定位
% ==================================================================

function seedPath = getSeedPath()
% 优先从函数所在目录查找种子文件（兼容 MATLAB Compiler 打包）
seedPath = '';
try
    funcPath = which('generateTemplate', '-all');
    if ischar(funcPath)
        funcDir = fileparts(funcPath);
    else
        % which 返回 cell 时取第一个
        funcDir = fileparts(funcPath{1});
    end
    candidate = fullfile(funcDir, 'AutoCreateSldd_TemplateSeed.xlsx');
    if exist(candidate, 'file')
        seedPath = candidate;
    end
end
end


function expectedPath = getExpectedSeedPath()
try
    funcPath = which('generateTemplate', '-all');
    if ischar(funcPath)
        funcDir = fileparts(funcPath);
    else
        funcDir = fileparts(funcPath{1});
    end
    expectedPath = fullfile(funcDir, 'AutoCreateSldd_TemplateSeed.xlsx');
catch
    expectedPath = 'AutoCreateSldd_TemplateSeed.xlsx';
end
end


% ==================================================================
%  通用辅助函数
% ==================================================================

function uniquePath = getUniqueFilePath(targetPath)
[pathStr, fileName, ext] = fileparts(targetPath);
uniquePath = targetPath;
index = 1;
while exist(uniquePath, 'file')
    uniquePath = fullfile(pathStr, sprintf('%s_%d%s', fileName, index, ext));
    index = index + 1;
end
end
