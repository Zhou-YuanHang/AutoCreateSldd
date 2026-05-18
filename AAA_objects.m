%% 自动生成的数据对象文件
%% 创建时间: 2026-05-17 02:08:21
%% 来源 Excel: AAA.xlsx
%%
%% 在 MATLAB 中运行此脚本即可在工作区创建所有 Simulink 对象

%% ================ Bus 对象 ================

%% ----- Bus: FaultBus -----
FaultBus = Simulink.Bus;
FaultBus.Description = '故障状态总线';
FaultBus.HeaderFile = 'fault_bus.h';
FaultBus.Alignment = -1;
FaultBus.PreserveElementDimensions = false;
FaultBus.DataScope = 'Auto';

%% ----- Bus: CtrlBus -----
CtrlBus = Simulink.Bus;
CtrlBus.Description = '控制总线';
CtrlBus.HeaderFile = 'ctrl_bus.h';
CtrlBus.Alignment = -1;
CtrlBus.PreserveElementDimensions = false;
CtrlBus.DataScope = 'Exported';


%% ================ Signal 对象 ================

%% ----- Signal: SignalA -----
SignalA = myPackage.Signal;
SignalA.CoderInfo.StorageClass = 'Custom';
SignalA.CoderInfo.CustomStorageClass = 'YS_SelectRAMSignal';
try SignalA.CoderInfo.CustomAttributes.HeaderFile = 'MCU_Fault_Signals.h'; catch, end
try SignalA.CoderInfo.CustomAttributes.DefinitionFile = 'MCU_Fault_Signals.c'; catch, end
SignalA.DataType = 'Bus: FaultBus';
SignalA.Description = '故障状态信号总线';
SignalA.DocUnits = '-';
SignalA.Dimensions = 1;
SignalA.Complexity = 'real';

%% ----- Signal: SignalB -----
SignalB = myPackage.Signal;
SignalB.CoderInfo.StorageClass = 'ExportedGlobal';
SignalB.DataType = 'single';
SignalB.Min = 0;
SignalB.Max = 100;
SignalB.Description = '位置反馈信号';
SignalB.DocUnits = 'm';
SignalB.Dimensions = 1;
SignalB.Complexity = 'real';


%% ================ Parameter 对象 ================

%% ----- Parameter: ParamA -----
ParamA = myPackage.Parameter;
ParamA.CoderInfo.StorageClass = 'Custom';
ParamA.CoderInfo.CustomStorageClass = 'YS_SelectRAMPara';
try ParamA.CoderInfo.CustomAttributes.HeaderFile = 'MCU_Fault_Paras.h'; catch, end
try ParamA.CoderInfo.CustomAttributes.DefinitionFile = 'MCU_Fault_Paras.c'; catch, end
ParamA.DataType = 'single';
ParamA.Dimensions = [1 1];
ParamA.Value = 600;
ParamA.Min = 0;
ParamA.Max = 1000;
ParamA.Description = 'A相电流故障限值设置';
ParamA.DocUnits = 'A';
ParamA.Complexity = 'real';

%% ----- Parameter: ParamB -----
ParamB = myPackage.Parameter;
ParamB.CoderInfo.StorageClass = 'ExportedGlobal';
ParamB.DataType = 'single';
ParamB.Dimensions = [1 1];
ParamB.Value = 0.35;
ParamB.Min = 0;
ParamB.Max = 10;
ParamB.Description = '速度环比例系数';
ParamB.DocUnits = '-';
ParamB.Complexity = 'real';

%% ----- Parameter: ParamBCopy -----
ParamBCopy = myPackage.Parameter;
ParamBCopy.CoderInfo.StorageClass = 'ExportedGlobal';
ParamBCopy.DataType = 'single';
ParamBCopy.Dimensions = [1 1];
ParamBCopy.Value = 0.35;
ParamBCopy.Min = 0;
ParamBCopy.Max = 10;
ParamBCopy.Description = '速度环比例系数';
ParamBCopy.DocUnits = '-';
ParamBCopy.Complexity = 'real';

%% ----- Parameter: ParamACopy -----
ParamACopy = myPackage.Parameter;
ParamACopy.CoderInfo.StorageClass = 'Custom';
ParamACopy.CoderInfo.CustomStorageClass = 'YS_SelectRAMPara';
try ParamACopy.CoderInfo.CustomAttributes.HeaderFile = 'MCU_Fault_Paras.h'; catch, end
try ParamACopy.CoderInfo.CustomAttributes.DefinitionFile = 'MCU_Fault_Paras.c'; catch, end
ParamACopy.DataType = 'single';
ParamACopy.Dimensions = [1 1];
ParamACopy.Value = 600;
ParamACopy.Min = 0;
ParamACopy.Max = 1000;
ParamACopy.Description = 'A相电流故障限值设置';
ParamACopy.DocUnits = 'A';
ParamACopy.Complexity = 'real';


%% ================ Const 对象 (Parameter + Define) ================

%% ----- Const: Mode -----
Mode = Simulink.Parameter;
Mode.CoderInfo.StorageClass = 'Custom';
Mode.CoderInfo.CustomStorageClass = 'Define';
Mode.DataType = 'uint8';
Mode.Value = 1;
Mode.CoderInfo.CustomAttributes.HeaderFile = 'Define.h';

%% ----- Const: ModeCopy -----
ModeCopy = Simulink.Parameter;
ModeCopy.CoderInfo.StorageClass = 'Custom';
ModeCopy.CoderInfo.CustomStorageClass = 'Define';
ModeCopy.DataType = 'uint8';
ModeCopy.Value = 1;
ModeCopy.CoderInfo.CustomAttributes.HeaderFile = 'Define.h';


%% ================ BusElement 定义 ================

%% ----- 为 Bus 'CtrlBus' 创建 2 个元素 -----
saveVarsTmp{1} = Simulink.BusElement;
saveVarsTmp{1}.Name = 'Position';
saveVarsTmp{1}.Description = '位置矢量';
saveVarsTmp{1}.DataType = 'double';
saveVarsTmp{1}.Dimensions = 1;
saveVarsTmp{1}.DocUnits = 'm';
saveVarsTmp{1}(2, 1) = Simulink.BusElement;
saveVarsTmp{1}(2, 1).Name = 'Velocity';
saveVarsTmp{1}(2, 1).Description = '速度矢量';
saveVarsTmp{1}(2, 1).DataType = 'double';
saveVarsTmp{1}(2, 1).Dimensions = 1;
saveVarsTmp{1}(2, 1).DocUnits = 'm/s';
CtrlBus.Elements = saveVarsTmp{1};
clear saveVarsTmp;

%% ----- 为 Bus 'FaultBus' 创建 3 个元素 -----
saveVarsTmp{1} = Simulink.BusElement;
saveVarsTmp{1}.Name = 'IcSensorFault';
saveVarsTmp{1}.Description = '电流传感器故障标志';
saveVarsTmp{1}.DataType = 'uint8';
saveVarsTmp{1}.Dimensions = 1;
saveVarsTmp{1}.DocUnits = '-';
saveVarsTmp{1}(2, 1) = Simulink.BusElement;
saveVarsTmp{1}(2, 1).Name = 'VoltageFault';
saveVarsTmp{1}(2, 1).Description = '母线电压故障标志';
saveVarsTmp{1}(2, 1).DataType = 'boolean';
saveVarsTmp{1}(2, 1).Dimensions = 1;
saveVarsTmp{1}(2, 1).DocUnits = '-';
saveVarsTmp{1}(3, 1) = Simulink.BusElement;
saveVarsTmp{1}(3, 1).Name = 'a';
saveVarsTmp{1}(3, 1).DataType = 'double';
saveVarsTmp{1}(3, 1).Dimensions = 1;
FaultBus.Elements = saveVarsTmp{1};
clear saveVarsTmp;

