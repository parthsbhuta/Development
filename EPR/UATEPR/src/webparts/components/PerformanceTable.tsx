import * as React from 'react';
import { Modal } from '@fluentui/react';
import { DataTable } from 'primereact/datatable';
import { Column } from 'primereact/column';
import "primeicons/primeicons.css";
import "primeflex/primeflex.css";
import "primereact/resources/themes/lara-light-indigo/theme.css";
import "primereact/resources/primereact.css";
import styles from '../css/Common.module.scss';

const EditIcon: any = require("../Images/Edit.png");
const SplitIcon: any = require("../Images/Split.png");
const CloseIcon: any = require("../Images/Close.png");
const checkIcon: any = require("../Images/check.png");
const DeleteIcon: any = require("../Images/Delete.png");
const StoppedIcon: any = require("../Images/Stopped.png");
const RunningIcon: any = require("../Images/Running.png");
const BothIcon: any = require("../Images/Both.png");

interface IPerformanceRow {
    id: number;
    byHoursId?: number;
    splitSourceId?: number;

    sortMinutes: number;

    Start: string;
    End: string;
    Time: string;
    MachineState: string;

    Class1: string;
    Class2: string;
    Class3: string;
    Class4: string;

    Class1Title: string;
    Class2Title: string;
    Class3Title: string;
    Class4Title: string;

    Remarks: string;
    IsSplit: string;
}

interface IPerformanceTableProps {
    runStopLogs: any[]; byHours: any[]; startTime: string; endTime: string;
    classTypes: any[]; classTypes2: any[]; classTypes3: any[]; classTypes4: any[];
    spService: any; SelectedStations: string; approvalErrorRowIds: number[]; eprId: number | null; eprDate: string; isEPRReadOnly: boolean;
    onByHoursRefresh?: () => Promise<void>; onRowsChange?: (rows: IPerformanceRow[]) => void;
}

const PerformanceTable: React.FC<IPerformanceTableProps> = ({
    runStopLogs, byHours, startTime, endTime,
    classTypes, classTypes2, classTypes3, classTypes4,
    spService, SelectedStations, approvalErrorRowIds, eprId, eprDate, isEPRReadOnly, onByHoursRefresh, onRowsChange,
}) => {

    const [currentTime, setCurrentTime] = React.useState<Date>(new Date());
    const [editingRowId, setEditingRowId] = React.useState<number | null>(null);
    const [editingRow, setEditingRow] = React.useState<IPerformanceRow | null>(null);
    const [showSplitPopup, setShowSplitPopup] = React.useState<boolean>(false);
    const [splitRow, setSplitRow] = React.useState<IPerformanceRow | null>(null);
    const [splitTime, setSplitTime] = React.useState<string>('');
    const [splitRows, setSplitRows] = React.useState<IPerformanceRow[]>([]);
    const [machineStateFilter, setMachineStateFilter] = React.useState<'Stopped' | 'Running' | 'Both'>('Stopped');

    /* Start of Utility Functions */
    const timeToMinutes = (time: string): number => {
        if (!time) {
            return 0;
        }
        const value = time.trim().toUpperCase();
        const match = value.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/);
        if (!match) {
            return 0;
        }

        let hours = Number(match[1]);
        const minutes = Number(match[2]);
        const period = match[3];

        if (period === 'AM' && hours === 12) {
            hours = 0;
        }
        if (period === 'PM' && hours !== 12) {
            hours += 12;
        }

        return (hours * 60) + minutes;
    };

    const getShiftDurationMinutes = (
        startTime: string,
        endTime: string
    ): number => {
        let start = timeToMinutes(startTime);
        let end = timeToMinutes(endTime);

        if (end <= start) {
            end += 1440;
        }

        return end - start;
    };

    const minutesToTime = (totalMinutes: number): string => {
        totalMinutes = ((totalMinutes % 1440) + 1440) % 1440;
        let hours = Math.floor(totalMinutes / 60);
        const minutes = totalMinutes % 60;
        const period = hours >= 12 ? 'PM' : 'AM';
        let displayHours = hours % 12;
        if (displayHours === 0) {
            displayHours = 12;
        }
        return (`${String(displayHours).padStart(2, '0')}:` + `${String(minutes).padStart(2, '0')} ` + period);
    };

    const parseRunStopDateTime = (value: string): Date | null => {
        if (!value) {
            return null;
        }
        const match = value.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})$/);
        if (!match) {
            return null;
        }

        const month = Number(match[1]);
        const day = Number(match[2]);
        const year = Number(match[3]);
        const hours = Number(match[4]);
        const minutes = Number(match[5]);
        const seconds = Number(match[6]);
        return new Date(year, month - 1, day, hours, minutes, seconds);
    };

    const getDuration = (start: number, end: number): string => {
        let difference = end - start;
        if (difference < 0) { difference += 1440; }
        const hours = Math.floor(difference / 60);
        const minutes = difference % 60;
        return (`${String(hours).padStart(2, '0')}:` + `${String(minutes).padStart(2, '0')}`);
    };

    const normalizeTime = (value: any): string => {
        if (!value) {
            return '';
        }

        const date = new Date(value);
        if (!isNaN(date.getTime()) && value.toString().includes('T')) {
            return minutesToTime(
                (date.getHours() * 60) + date.getMinutes()
            );
        }

        const valueString = String(value).trim();
        const match = valueString.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
        if (match) {
            let hours = Number(match[1]);
            const minutes = Number(match[2]);
            const period = match[3].toUpperCase();

            if (period === 'AM' && hours === 12) {
                hours = 0;
            }

            if (period === 'PM' && hours !== 12) {
                hours += 12;
            }
            return minutesToTime((hours * 60) + minutes);
        }

        const simpleMatch = valueString.match(/^(\d{1,2}):(\d{2})$/);
        if (simpleMatch) {
            return minutesToTime((Number(simpleMatch[1]) * 60) + Number(simpleMatch[2]));
        }
        return valueString;
    };

    const isLastTimelineRow = (row: IPerformanceRow): boolean => {
        if (!row || !endTime) {
            return false;
        }

        const rowEnd = normalizeTime(row.End);
        const eprEnd = normalizeTime(endTime);

        return rowEnd === eprEnd;
    };

    const isGrayLockedLastRow = (row: IPerformanceRow): boolean => {
        if (!isLastTimelineRow(row)) {
            return false;
        }

        if (!eprDate || !endTime) {
            return false;
        }

        const [year, month, day] = eprDate.split('-').map(Number);
        const shiftEndMinutes = timeToMinutes(endTime);

        if (shiftEndMinutes < 0) {
            return false;
        }

        const shiftEndDateTime = new Date(year, month - 1, day, Math.floor(shiftEndMinutes / 60), shiftEndMinutes % 60, 0);
        const startMinutes = timeToMinutes(startTime);

        if (startMinutes >= 0 && shiftEndMinutes <= startMinutes) {
            shiftEndDateTime.setDate(shiftEndDateTime.getDate() + 1);
        }

        const selectedDate = new Date(year, month - 1, day);
        const today = new Date();
        const selectedDateOnly = new Date(today.getFullYear(), today.getMonth(), today.getDate());
        const eprDateOnly = new Date(selectedDate.getFullYear(), selectedDate.getMonth(), selectedDate.getDate());

        if (eprDateOnly.getTime() !== selectedDateOnly.getTime()) {
            return false;
        }

        return currentTime.getTime() < shiftEndDateTime.getTime();
    };

    const getRunStopFieldMapping = (
        station: string
    ): {
        timestampField: string;
        dateField: string;
        machineStateField: string;
    } => {
        const normalizedStation = String(station || '')
            .replace(/\s+/g, '')
            .toLowerCase();

        switch (normalizedStation) {
            case 'heattreat':
            case 'erwmill':
                return {
                    timestampField: 'field_4',
                    dateField: 'field_10',
                    machineStateField: 'field_5'
                };

            case 'hydro':
            case 'ndt':
            case 'threadline1':
            case 'threadline2':
                return {
                    timestampField: 'field_1',
                    dateField: 'field_3',
                    machineStateField: 'field_4'
                };

            default:
                return {
                    timestampField: 'field_4',
                    dateField: 'field_10',
                    machineStateField: 'field_5'
                };
        }
    };

    React.useEffect(() => {
        const timer = window.setInterval(() => { setCurrentTime(new Date()); }, 30000);
        return () => { window.clearInterval(timer); };
    }, []);
    /* End of Utility Functions */

    const tableData = React.useMemo<IPerformanceRow[]>(() => {
        if (!startTime || !endTime || !eprDate) {
            return [];
        }

        const [year, month, day] = eprDate.split('-').map(Number);

        const startTotalMinutes = timeToMinutes(startTime);
        const endTotalMinutes = timeToMinutes(endTime);

        const eprStartDateTime = new Date(year, month - 1, day, Math.floor(startTotalMinutes / 60), startTotalMinutes % 60, 0);
        const eprEndDateTime = new Date(year, month - 1, day, Math.floor(endTotalMinutes / 60), endTotalMinutes % 60, 0);

        if (eprEndDateTime <= eprStartDateTime) {
            eprEndDateTime.setDate(eprEndDateTime.getDate() + 1);
        }

        const now = new Date();
        const eprEndHasPassed = now.getTime() >= eprEndDateTime.getTime();

        // const logs = (runStopLogs || []).filter((log: any) => {
        //     if (!log.field_4) { return false; }
        //     return parseRunStopDateTime(log.field_4) !== null;
        // }).map((log: any) => ({ ...log, dateTime: parseRunStopDateTime(log.field_4) }))
        //     .filter((log: any) => log.dateTime !== null).sort((a: any, b: any) => a.dateTime.getTime() - b.dateTime.getTime());

        const { timestampField, machineStateField } = getRunStopFieldMapping(SelectedStations);

        const logs = (runStopLogs || []).filter((log: any) => {
            if (!log[timestampField]) {
                return false;
            }

            return parseRunStopDateTime(log[timestampField]) !== null;
        }).map((log: any) => ({
            ...log,
            dateTime: parseRunStopDateTime(log[timestampField])
        }))
            .filter((log: any) => log.dateTime !== null)
            .sort((a: any, b: any) => a.dateTime.getTime() - b.dateTime.getTime());

        let initialMachineState = 'Stopped';

        const logsBeforeStart = logs.filter((log: any) => {
            return (log.dateTime && log.dateTime.getTime() <= eprStartDateTime.getTime());
        }).sort((a: any, b: any) => b.dateTime.getTime() - a.dateTime.getTime());

        if (logsBeforeStart.length > 0) {
            initialMachineState = logsBeforeStart[0][machineStateField] || 'Stopped';
        }

        const logsInsideEPR = logs.filter((log: any) => {
            if (!log.dateTime) {
                return false;
            }
            const logTime = log.dateTime.getTime();
            return (logTime > eprStartDateTime.getTime() && logTime < eprEndDateTime.getTime());
        });

        const timeToDateTime = (time: string): Date | null => {
            const minutes = timeToMinutes(time);
            if (minutes < 0) {
                return null;
            }

            let offsetMinutes = minutes - startTotalMinutes;
            if (offsetMinutes < 0) { offsetMinutes += 1440; }
            return new Date(eprStartDateTime.getTime() + offsetMinutes * 60 * 1000);
        };

        const savedByHours = (byHours || [])
            .map((item: any) => {
                const savedStart = normalizeTime(item.TimeStart);
                const savedEnd = normalizeTime(item.TimeEnd);
                const savedStartDateTime = timeToDateTime(savedStart);
                const savedEndDateTime = timeToDateTime(savedEnd);

                if (!savedStartDateTime || !savedEndDateTime) {
                    return null;
                }

                let adjustedEndDateTime = new Date(savedEndDateTime);
                if (adjustedEndDateTime <= savedStartDateTime) {
                    adjustedEndDateTime.setDate(adjustedEndDateTime.getDate() + 1);
                }
                return { item, startDateTime: savedStartDateTime, endDateTime: adjustedEndDateTime };
            }).filter((item: any): item is { item: any; startDateTime: Date; endDateTime: Date; } => item !== null).filter((saved: any) => {
                return (saved.endDateTime.getTime() > eprStartDateTime.getTime() && saved.startDateTime.getTime() < eprEndDateTime.getTime());
            }).sort((a: any, b: any) => a.startDateTime.getTime() - b.startDateTime.getTime());
        const boundaryMap = new Map<number, Date>();

        const addBoundary = (dateTime: Date): void => {
            const time = dateTime.getTime();
            if (time >= eprStartDateTime.getTime() && time <= eprEndDateTime.getTime()) {
                boundaryMap.set(time, new Date(dateTime));
            }
        };

        addBoundary(eprStartDateTime);
        addBoundary(eprEndDateTime);

        logsInsideEPR.forEach((log: any) => {
            if (log.dateTime) {
                addBoundary(log.dateTime);
            }
        });

        savedByHours.forEach((saved: any) => {
            addBoundary(saved.startDateTime);
            addBoundary(saved.endDateTime);
        });

        const boundaries = Array.from(boundaryMap.values()).sort((a, b) => a.getTime() - b.getTime());

        if (logs.length === 0) {
            if (savedByHours.length > 0) {
                return savedByHours.map((saved: any, index: number) => {
                    const item = saved.item;
                    const rowStartDateTime = saved.startDateTime;
                    const rowEndDateTime = saved.endDateTime;
                    return {
                        id: Number(item.Id) || index + 1,
                        byHoursId: item.Id ? Number(item.Id) : undefined,
                        sortMinutes: rowStartDateTime.getTime(),
                        Start: normalizeTime(item.TimeStart),
                        End: normalizeTime(item.TimeEnd),
                        Time: item.TimeTotal || getDuration(rowStartDateTime.getHours() * 60 + rowStartDateTime.getMinutes(), rowEndDateTime.getHours() * 60 + rowEndDateTime.getMinutes()),
                        MachineState: item.MachineState || 'Stopped',
                        Class1: item.ClassTypeId ? String(item.ClassTypeId) : '',
                        Class2: item.ClassId ? String(item.ClassId) : '',
                        Class3: item.VoiceTypeId ? String(item.VoiceTypeId) : '',
                        Class4: item.VoiceId ? String(item.VoiceId) : '',
                        Class1Title: item.ClassType?.ViewTitle || '',
                        Class2Title: item.Class?.ViewTitle || '',
                        Class3Title: item.VoiceType?.ViewTitle || '',
                        Class4Title: item.Voice?.ViewTitle || '',
                        Remarks: item.Remarks || '',
                        IsSplit: item.IsSplit || ''
                    };
                }).sort((a, b) => a.sortMinutes - b.sortMinutes);
            }

            if (!eprEndHasPassed) {
                return [];
            }

            return [
                {
                    id: Date.now(),
                    byHoursId: undefined,
                    sortMinutes: eprStartDateTime.getTime(),
                    Start: minutesToTime(eprStartDateTime.getHours() * 60 + eprStartDateTime.getMinutes()),
                    End: minutesToTime(eprEndDateTime.getHours() * 60 + eprEndDateTime.getMinutes()),
                    Time: getDuration(eprStartDateTime.getHours() * 60 + eprStartDateTime.getMinutes(), eprEndDateTime.getHours() * 60 + eprEndDateTime.getMinutes()),
                    MachineState: 'Stopped',
                    Class1: '',
                    Class2: '',
                    Class3: '',
                    Class4: '',
                    Class1Title: '',
                    Class2Title: '',
                    Class3Title: '',
                    Class4Title: '',
                    Remarks: '',
                    IsSplit: ''
                }
            ];
        }

        const rows: IPerformanceRow[] = [];
        for (let index = 0; index < boundaries.length - 1; index++) {
            const rowStartDateTime = boundaries[index];
            const rowEndDateTime = boundaries[index + 1];
            if (rowEndDateTime.getTime() <= rowStartDateTime.getTime()) {
                continue;
            }

            const stateLogs = logs.filter((log: any) => {
                return (log.dateTime && log.dateTime.getTime() <= rowStartDateTime.getTime());
            });

            let machineState = initialMachineState;
            if (stateLogs.length > 0) {
                const latestLog = stateLogs[stateLogs.length - 1];
                machineState = latestLog[machineStateField] || machineState;
            }

            const matchingByHour = savedByHours.find((saved: any) => {
                return (saved.startDateTime.getTime() <= rowStartDateTime.getTime() && saved.endDateTime.getTime() >= rowEndDateTime.getTime());
            });

            const isSplitContinuation = savedByHours.some((saved: any) => saved.endDateTime.getTime() === rowStartDateTime.getTime());
            const rowStartMinutes = rowStartDateTime.getHours() * 60 + rowStartDateTime.getMinutes();
            const rowEndMinutes = rowEndDateTime.getHours() * 60 + rowEndDateTime.getMinutes();
            rows.push({
                id: Number(matchingByHour?.item?.Id) || Number(logs.find((log: any) => log.dateTime?.getTime() === rowEndDateTime.getTime())?.Id) || Date.now() + index,
                byHoursId: matchingByHour?.item?.Id ? Number(matchingByHour.item.Id) : undefined,
                sortMinutes: rowStartDateTime.getTime(),
                Start: minutesToTime(rowStartMinutes),
                End: minutesToTime(rowEndMinutes),
                Time: matchingByHour?.item?.TimeTotal || getDuration(rowStartMinutes, rowEndMinutes),
                MachineState: matchingByHour?.item?.MachineState || machineState,
                Class1: matchingByHour?.item?.ClassTypeId ? String(matchingByHour.item.ClassTypeId) : '',
                Class2: matchingByHour?.item?.ClassId ? String(matchingByHour.item.ClassId) : '',
                Class3: matchingByHour?.item?.VoiceTypeId ? String(matchingByHour.item.VoiceTypeId) : '',
                Class4: matchingByHour?.item?.VoiceId ? String(matchingByHour.item.VoiceId) : '',
                Class1Title: matchingByHour?.item?.ClassType?.ViewTitle || '',
                Class2Title: matchingByHour?.item?.Class?.ViewTitle || '',
                Class3Title: matchingByHour?.item?.VoiceType?.ViewTitle || '',
                Class4Title: matchingByHour?.item?.Voice?.ViewTitle || '',
                Remarks: matchingByHour?.item?.Remarks || '',
                IsSplit: matchingByHour?.item?.IsSplit || (isSplitContinuation ? 'Yes' : '')
            });
        }

        return rows;
    }, [
        runStopLogs,
        byHours,
        startTime,
        endTime,
        eprDate,
        SelectedStations
    ]);

    const displayTableData = React.useMemo<IPerformanceRow[]>(() => {
        const getRowRange = (row: IPerformanceRow): { start: number; end: number } => {
            let start = timeToMinutes(row.Start);
            let end = timeToMinutes(row.End);
            const eprStartMinutes = timeToMinutes(startTime);
            if (start < eprStartMinutes) {
                start += 1440;
            }

            if (end < eprStartMinutes) {
                end += 1440;
            }

            if (end <= start) {
                end += 1440;
            }

            return { start, end };
        };

        const hasTemporarySplitCoveringRow = (row: IPerformanceRow): boolean => {
            if (splitRows.length === 0) {
                return false;
            }

            const rowRange = getRowRange(row);
            const splitRanges = splitRows.map(splitRow => ({ row: splitRow, range: getRowRange(splitRow) })).filter(({ range }) => {
                return (range.start < rowRange.end && range.end > rowRange.start);
            });
            if (splitRanges.length === 0) {
                return false;
            }

            const coveredMinutes = splitRanges.sort((a, b) => a.range.start - b.range.start).reduce((currentEnd, item) => {
                if (item.range.start > currentEnd) {
                    return currentEnd;
                }
                return Math.max(currentEnd, item.range.end);
            },
                rowRange.start
            );
            return (coveredMinutes >= rowRange.end);
        };

        const result: IPerformanceRow[] = [];

        tableData.forEach((row: IPerformanceRow) => {
            if (!hasTemporarySplitCoveringRow(row)) {
                if (row.Start !== row.End) {
                    result.push(row);
                }
            }
        });

        splitRows.forEach((splitRow: IPerformanceRow) => {
            if (splitRow.Start !== splitRow.End) {
                result.push(splitRow);
            }
        });

        const uniqueRows = new Map<string, IPerformanceRow>();
        result.forEach((row: IPerformanceRow) => {
            const key = `${row.Start}|${row.End}|${row.byHoursId || 'temp'}`;
            if (!uniqueRows.has(key) || row.byHoursId) {
                uniqueRows.set(key, row);
            }
        });

        return Array.from(uniqueRows.values()).sort((a, b) => b.sortMinutes - a.sortMinutes);
    }, [tableData, splitRows, startTime]);

    React.useEffect(() => {
        if (onRowsChange) {
            onRowsChange(displayTableData);
        }
    }, [displayTableData, onRowsChange]);

    const handleEdit = (row: IPerformanceRow): void => {
        if (isEPRReadOnly) {
            return;
        }
        if (isGrayLockedLastRow(row)) {
            return;
        }
        setEditingRowId(row.id);
        setEditingRow({ ...row, Class1: row.Class1 || '', Class2: row.Class2 || '', Class3: row.Class3 || '', Class4: row.Class4 || '' });
    };

    const handleCancel = (): void => {
        setEditingRowId(null);
        setEditingRow(null);
    };

    const handleFieldChange = (field: keyof IPerformanceRow, value: string): void => {
        if (!editingRow) {
            return;
        }
        setEditingRow({ ...editingRow, [field]: value });
    };

    const handleSave = async (): Promise<void> => {
        if (!editingRow) {
            return;
        }

        if (isGrayLockedLastRow(editingRow)) {
            setEditingRowId(null);
            setEditingRow(null);
            alert('This performance row is locked until the shift end time.');
            return;
        }

        if (!eprId) {
            console.error('EPR ID is missing.');
            return;
        }

        const startDateTime = new Date(`${eprDate} ${editingRow.Start}`);
        const endDateTime = new Date(`${eprDate} ${editingRow.End}`);
        if (endDateTime <= startDateTime) {
            endDateTime.setDate(endDateTime.getDate() + 1);
        }

        try {
            const payload = {
                EPR_ID: eprId,
                StartDate: startDateTime,
                EndDate: endDateTime,
                TimeStart: editingRow.Start,
                TimeEnd: editingRow.End,
                TimeTotal: editingRow.Time,
                MachineState: editingRow.MachineState,
                ClassTypeId: editingRow.Class1 ? Number(editingRow.Class1) : null,
                ClassId: editingRow.Class2 ? Number(editingRow.Class2) : null,
                VoiceTypeId: editingRow.Class3 ? Number(editingRow.Class3) : null,
                VoiceId: editingRow.Class4 ? Number(editingRow.Class4) : null,
                Remarks: editingRow.Remarks,
                IsSplit: editingRow.IsSplit
            };

            if (editingRow.byHoursId) {
                await spService.updateEPRByHours(SelectedStations, editingRow.byHoursId, payload);
            }
            else {
                await spService.addEPRByHours(SelectedStations, payload);
            }

            if (onByHoursRefresh) {
                await onByHoursRefresh();
            }

            if (editingRow.splitSourceId !== undefined) {
                setSplitRows(prev =>
                    prev.filter(
                        row => row.id !== editingRow.id
                    )
                );
            }

            setEditingRowId(null);
            setEditingRow(null);
        }
        catch (error) {
            console.error('Error saving EPR By Hours:', error);
        }
    };

    /* Start of DataTable Column Design */
    const renderInput = (field: keyof IPerformanceRow): JSX.Element => {
        return (<input type="text" className={styles.performance_table_input} value={String(editingRow?.[field] ?? '')} onChange={(e) => handleFieldChange(field, e.target.value)} />);
    };

    const renderStart = (row: IPerformanceRow): JSX.Element | string => {
        if (editingRowId !== row.id) {
            return row.Start;
        }
        return (<input type="text" className={styles.performance_table_input} value={editingRow?.Start || ''} disabled readOnly />);
    };

    const renderEnd = (row: IPerformanceRow): JSX.Element | string => {
        if (editingRowId !== row.id) {
            return row.End;
        }
        return (<input type="text" className={styles.performance_table_input} value={editingRow?.End || ''} disabled readOnly />);
    };

    const renderTime = (row: IPerformanceRow): JSX.Element | string => {
        if (editingRowId !== row.id) {
            return row.Time;
        }
        return (<input type="text" className={styles.performance_table_input} value={row.Time || ''} disabled readOnly />);
    };

    const renderMachineState = (row: IPerformanceRow): JSX.Element | string => {
        if (editingRowId !== row.id) {
            return (
                <span className={row.MachineState === 'Running' ? styles.machine_state_running : styles.machine_state_stopped}>
                    {row.MachineState}
                </span>
            );
        }
        return (
            <input type="text" className={styles.performance_table_input} value={editingRow?.MachineState || ''} disabled readOnly />
        );
    };

    const filteredClass2Options = React.useMemo(() => {
        if (!editingRow?.Class1) {
            return [];
        }

        return (classTypes2 || []).filter(
            (item: any) =>
                Number(item.ClassTypeId) === Number(editingRow.Class1)
        );
    }, [classTypes2, editingRow?.Class1]);

    const filteredClass3Options = React.useMemo(() => {
        if (!editingRow?.Class2) {
            return [];
        }

        return (classTypes3 || []).filter(
            (item: any) =>
                Number(item.ClassId) === Number(editingRow.Class2)
        );
    }, [classTypes3, editingRow?.Class2]);

    const filteredClass4Options = React.useMemo(() => {
        if (!editingRow?.Class3) {
            return [];
        }

        return (classTypes4 || []).filter((item: any) => Number(item.VoiceTypeId) === Number(editingRow.Class3));
    }, [classTypes4, editingRow?.Class3]);

    const renderClass1 = (row: IPerformanceRow): JSX.Element | string => {
        if (editingRowId !== row.id) {
            return (
                <span className={isApprovalFieldMissing(row, 'Class1') ? styles.approval_missing_field : ''}>
                    {row.Class1Title || ''}
                </span>
            );
        }

        return (
            <select className={styles.performance_table_input} value={editingRow?.Class1 || ''}
                onChange={(e) => {
                    const value = e.target.value;
                    setEditingRow(prev => {
                        if (!prev) {
                            return prev;
                        }
                        return { ...prev, Class1: value, Class2: '', Class3: '', Class4: '' };
                    });
                }}>
                <option value="">Select Class</option>
                {classTypes.map((item: any, index: number) => (
                    <option key={item.Id || index} value={String(item.Id)}>
                        {item.ViewTitle}
                    </option>
                ))}
            </select>
        );
    };

    const renderClass2 = (row: IPerformanceRow): JSX.Element | string => {
        if (editingRowId !== row.id) {
            return (
                <span className={isApprovalFieldMissing(row, 'Class2') ? styles.approval_missing_field : ''}>
                    {row.Class2Title || ''}
                </span>
            );
        }

        return (
            <select className={styles.performance_table_input} value={editingRow?.Class2 || ''}
                onChange={(e) => {
                    const value = e.target.value;
                    setEditingRow(prev => {
                        if (!prev) {
                            return prev;
                        }
                        return { ...prev, Class2: value, Class3: '', Class4: '' };
                    });
                }}>
                <option value="">Select Class</option>
                {filteredClass2Options.map((item: any, index: number) => (
                    <option key={item.Id || index} value={String(item.Id)}>
                        {item.ViewTitle}
                    </option>
                ))}
            </select>
        );
    };

    const renderClass3 = (row: IPerformanceRow): JSX.Element | string => {
        if (editingRowId !== row.id) {
            return (
                <span className={isApprovalFieldMissing(row, 'Class3') ? styles.approval_missing_field : ''}>
                    {row.Class3Title || ''}
                </span>
            );
        }

        return (
            <select className={styles.performance_table_input} value={editingRow?.Class3 || ''}
                onChange={(e) => {
                    const value = e.target.value;
                    setEditingRow(prev => {
                        if (!prev) {
                            return prev;
                        }
                        return { ...prev, Class3: value, Class4: '' };
                    });
                }}>
                <option value="">Select Class</option>
                {filteredClass3Options.map((item: any, index: number) => (
                    <option key={item.Id || index} value={String(item.Id)}>
                        {item.ViewTitle}
                    </option>
                ))}
            </select>
        );
    };

    const renderClass4 = (row: IPerformanceRow): JSX.Element | string => {
        if (editingRowId !== row.id) {
            return row.Class4Title || '';
        }

        return (
            <select className={styles.performance_table_input} value={editingRow?.Class4 || ''}
                onChange={(e) => {
                    setEditingRow(prev => {
                        if (!prev) {
                            return prev;
                        }
                        return { ...prev, Class4: e.target.value };
                    });
                }}>
                <option value="">Select Class</option>
                {filteredClass4Options.map((item: any, index: number) => (
                    <option key={item.Id || index} value={String(item.Id)}>
                        {item.ViewTitle}
                    </option>
                ))}
            </select>
        );
    };

    const renderRemarks = (row: IPerformanceRow): JSX.Element | string => {
        if (editingRowId !== row.id) {
            return row.Remarks;
        }
        return renderInput('Remarks');
    };

    const renderAction = (row: IPerformanceRow): JSX.Element => {
        const isGrayLocked = isGrayLockedLastRow(row);
        if (editingRowId === row.id) {
            return (
                <div className={styles.actionBtnBox}>
                    <button type="button" className={styles.table_btn} onClick={handleSave} title="Save">
                        <img src={checkIcon} />
                    </button>
                    <button type="button" className={styles.table_btn} onClick={handleCancel} title="Cancel">
                        <img src={CloseIcon} />
                    </button>
                </div>
            );
        }

        if (row.MachineState === 'Running' || isGrayLocked) {
            return <></>;
        }

        return (
            <>
                {!isEPRReadOnly && (
                    <div className={styles.actionBtnBox}>
                        <button type="button" className={styles.table_btn} onClick={() => handleEdit(row)} title="Edit">
                            <img src={EditIcon} />
                        </button>
                        {row.byHoursId && (
                            <button type="button" className={styles.table_btn} onClick={() => handleDelete(row)} title="Delete">
                                <img src={DeleteIcon} />
                            </button>
                        )}
                    </div>
                )}
            </>
        );
    };
    /* End of DataTable Column Design */

    const handleSplitClick = (row: IPerformanceRow): void => {

        if (isEPRReadOnly) {
            return;
        }

        if (isGrayLockedLastRow(row)) {
            return;
        }

        const isUnsavedSplit = splitRows.some(splitRow => splitRow.id === row.id);

        if (row.byHoursId && !isUnsavedSplit) {
            return;
        }

        setSplitRow(row);
        setSplitTime('');
        setShowSplitPopup(true);
    };

    const validateSplitTime = (row: IPerformanceRow, value: string): boolean => {
        const splitMinutes = timeToMinutes(value);
        const startMinutes = timeToMinutes(row.Start);
        let endMinutes = timeToMinutes(row.End);

        if (splitMinutes < 0 || startMinutes < 0 || endMinutes < 0) {
            return false;
        }

        if (endMinutes <= startMinutes) {
            endMinutes += 1440;
        }

        let normalizedSplitMinutes = splitMinutes;

        if (normalizedSplitMinutes < startMinutes) {
            normalizedSplitMinutes += 1440;
        }

        if (normalizedSplitMinutes <= startMinutes) {
            return false;
        }

        if (normalizedSplitMinutes >= endMinutes) {
            return false;
        }

        return true;
    };

    const confirmSplit = (): void => {
        if (!splitRow) {
            return;
        }

        if (!validateSplitTime(splitRow, splitTime)) {
            alert(
                `Split time must be between ${splitRow.Start} and ${splitRow.End}.`
            );
            return;
        }

        const startMinutes = timeToMinutes(splitRow.Start);
        let endMinutes = timeToMinutes(splitRow.End);
        let splitMinutes = timeToMinutes(splitTime);

        if (endMinutes <= startMinutes) {
            endMinutes += 1440;
        }

        if (splitMinutes < startMinutes) {
            splitMinutes += 1440;
        }

        const rootSplitSourceId = splitRow.splitSourceId !== undefined ? splitRow.splitSourceId : splitRow.id;
        const originalSortTime = splitRow.sortMinutes;
        const firstRow: IPerformanceRow = {
            ...splitRow,
            id: Date.now(),
            byHoursId: undefined,
            splitSourceId: rootSplitSourceId,
            sortMinutes: originalSortTime,
            Start: minutesToTime(startMinutes),
            End: minutesToTime(splitMinutes),
            Time: getDuration(startMinutes, splitMinutes),
            IsSplit: 'Yes'
        };

        const secondRow: IPerformanceRow = {
            ...splitRow,
            id: Date.now() + 1,
            byHoursId: undefined,
            splitSourceId: rootSplitSourceId,
            sortMinutes: originalSortTime + ((splitMinutes - startMinutes) * 60 * 1000),
            Start: minutesToTime(splitMinutes),
            End: minutesToTime(endMinutes),
            Time: getDuration(splitMinutes, endMinutes),
            Class1: '',
            Class2: '',
            Class3: '',
            Class4: '',
            Class1Title: '',
            Class2Title: '',
            Class3Title: '',
            Class4Title: '',
            Remarks: '',
            IsSplit: 'Yes'
        };

        setSplitRows(prev => [...prev.filter(row => row.id !== splitRow.id), firstRow, secondRow]);
        setShowSplitPopup(false);
        setSplitRow(null);
        setSplitTime('');
    };

    const handleDelete = async (row: IPerformanceRow): Promise<void> => {
        if (!row.byHoursId) {
            return;
        }
        const confirmed = window.confirm('Are you sure you want to delete this performance entry?');
        if (!confirmed) {
            return;
        }
        try {
            await spService.deleteEPRByHours(SelectedStations, row.byHoursId);
            if (onByHoursRefresh) {
                await onByHoursRefresh();
            }
        }
        catch (error) {
            console.error('Error deleting EPR By Hours:', error);
        }
    };

    const timeOptions = React.useMemo((): string[] => {
        const times: string[] = [];
        for (let hour = 0; hour < 24; hour++) {
            for (let minute = 0; minute < 60; minute += 5) {
                const period = hour >= 12 ? 'PM' : 'AM';
                const displayHour = hour % 12 === 0 ? 12 : hour % 12;
                const displayMinute = minute.toString().padStart(2, '0');
                times.push(`${displayHour}:${displayMinute} ${period}`);
            }
        }
        return times;
    }, []);

    const filteredTableData = React.useMemo(() => {
        if (displayTableData.length <= 2) {
            return displayTableData;
        }
        const firstRow = displayTableData[displayTableData.length - 1];
        const lastRow = displayTableData[0];
        if (machineStateFilter === 'Both') {
            return displayTableData;
        }
        const filteredRows = displayTableData.filter((row: IPerformanceRow) => row.MachineState === machineStateFilter);
        const result = [...filteredRows];
        if (!result.some((row) => row.id === firstRow.id)) {
            result.push(firstRow);
        }
        if (lastRow.id !== firstRow.id && !result.some((row) => row.id === lastRow.id)) {
            result.push(lastRow);
        }
        return result.sort((a, b) => b.sortMinutes - a.sortMinutes);
    }, [displayTableData, machineStateFilter]);

    const isApprovalFieldMissing = (row: IPerformanceRow, field: 'Class1' | 'Class2' | 'Class3'): boolean => {
        if (row.MachineState !== 'Stopped') {
            return false;
        }
        if (approvalErrorRowIds.indexOf(row.id) === -1) {
            return false;
        }
        return !row[field];
    };

    const performanceSummary = React.useMemo(() => {
        const parseDuration = (value: string): number => {
            if (!value) {
                return 0;
            }
            const parts = value.split(':').map(Number);
            if (parts.length !== 2) {
                return 0;
            }
            return (parts[0] * 60) + parts[1];
        };

        let totalMinutes = 0;
        let stoppedMinutes = 0;
        let runningMinutes = 0;

        displayTableData.forEach((row: IPerformanceRow) => {
            const minutes = parseDuration(row.Time);
            totalMinutes += minutes;
            if (row.MachineState === 'Stopped') {
                stoppedMinutes += minutes;
            }
            if (row.MachineState === 'Running') {
                runningMinutes += minutes;
            }
        });

        const formatDuration = (minutes: number): string => {
            const hours = Math.floor(minutes / 60);
            const mins = minutes % 60;
            return `${String(hours).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
        };

        const shiftDuration = getShiftDurationMinutes(startTime, endTime);
        return {
            totalTime: formatDuration(totalMinutes),
            stoppedTime: formatDuration(stoppedMinutes),
            runningTime: formatDuration(runningMinutes),
            shiftDuration: formatDuration(shiftDuration)
        };

    }, [displayTableData, startTime, endTime]);

    const sortTime = (event: any): void => {
        event.data = [...event.data].sort((a: IPerformanceRow, b: IPerformanceRow) => {
            const aTime = timeToMinutes(a[event.field as keyof IPerformanceRow] as string);
            const bTime = timeToMinutes(b[event.field as keyof IPerformanceRow] as string);
            return (aTime - bTime) * event.order;
        });
    };

    const sortDuration = (event: any): void => {
        event.data = [...event.data].sort(
            (a: IPerformanceRow, b: IPerformanceRow) => {
                const [aHours, aMinutes] = String(a.Time || '00:00').split(':').map(Number);
                const [bHours, bMinutes] = String(b.Time || '00:00').split(':').map(Number);
                const aTotal = (aHours * 60) + aMinutes;
                const bTotal = (bHours * 60) + bMinutes;
                return (aTotal - bTotal) * event.order;
            }
        );
    };

    const sortText = (field: keyof IPerformanceRow) => (event: any): void => {
        event.data = [...event.data].sort(
            (a: IPerformanceRow, b: IPerformanceRow) => {
                const aValue = String(a[field] || '').toLowerCase();
                const bValue = String(b[field] || '').toLowerCase();
                return aValue.localeCompare(bValue) * event.order;
            }
        );
    };



    return (
        <>
            <div className={styles.filter_export_row}>
                <span className={styles.msf_title}>Machine State</span>
                <ul className={styles.machine_state_filter}>
                    <li className={machineStateFilter === 'Stopped' ? styles.active : ''}>
                        <label>
                            <img src={StoppedIcon} />
                            <input type="radio" name="machineState" value="Stopped" checked={machineStateFilter === 'Stopped'} onChange={() => setMachineStateFilter('Stopped')} />
                            <span>Stopped</span>
                            <strong>({performanceSummary.stoppedTime})</strong>
                        </label>
                    </li>

                    <li className={machineStateFilter === 'Running' ? styles.active : ''}>
                        <label>
                            <img src={RunningIcon} />
                            <input type="radio" name="machineState" value="Running" checked={machineStateFilter === 'Running'} onChange={() => setMachineStateFilter('Running')} />
                            <span>Running</span>
                            <strong>({performanceSummary.runningTime})</strong>
                        </label>
                    </li>

                    <li className={machineStateFilter === 'Both' ? styles.active : ''}>
                        <label>
                            <img src={BothIcon} />
                            <input type="radio" name="machineState" value="Both" checked={machineStateFilter === 'Both'} onChange={() => setMachineStateFilter('Both')} />
                            <span>Both</span>
                            <strong>({performanceSummary.totalTime})</strong>
                        </label>
                    </li>
                </ul>
            </div>

            <div className={styles.custom_table}>
                <DataTable value={filteredTableData} paginator rows={10} responsiveLayout="scroll" emptyMessage="No Data Found"
                    rowClassName={(rowData: IPerformanceRow) => {
                        const isUnsavedSplit = splitRows.some(splitRow => splitRow.id === rowData.id);
                        const isGrayLocked = isGrayLockedLastRow(rowData);
                        if (isGrayLocked) {
                            return styles.gray_locked_row;
                        }

                        if (isUnsavedSplit) {
                            return styles.split_row;
                        }

                        return '';
                    }}>
                    <Column field="Start" header="Start" body={renderStart} sortable sortFunction={sortTime} />
                    <Column field="End" header="End" body={renderEnd} sortable sortFunction={sortTime} />
                    <Column field="Time" header="Time" body={renderTime} sortable sortFunction={sortDuration} />
                    <Column field="MachineState" header="Machine State" body={renderMachineState} sortable sortFunction={sortText('MachineState')} />
                    <Column field="Class1Title" header="Class 1" body={renderClass1} sortable sortFunction={sortText('Class1Title')} />
                    <Column field="Class2Title" header="Class 2" body={renderClass2} sortable sortFunction={sortText('Class2Title')} />
                    <Column field="Class3Title" header="Class 3" body={renderClass3} sortable sortFunction={sortText('Class3Title')} />
                    <Column field="Class4Title" header="Class 4" body={renderClass4} sortable sortFunction={sortText('Class4Title')} />
                    <Column field="Remarks" header="Remarks" body={renderRemarks} />
                    <Column field="IsSplit" header="Is Split"
                        body={(rowData: IPerformanceRow) => {
                            if (rowData.IsSplit !== 'Yes') {
                                return null;
                            }
                            const isUnsavedSplit = splitRows.some(splitRow => splitRow.id === rowData.id);
                            return (
                                <div>
                                    <span>Yes</span>
                                    {isUnsavedSplit && (<span className={styles.unsavedSplitMark}>*</span>)}
                                </div>
                            );
                        }}
                    />
                    <Column header="Split" body={(rowData: IPerformanceRow) => {
                        const isGrayLocked = isGrayLockedLastRow(rowData);
                        if (rowData.byHoursId || rowData.MachineState === 'Running' || isGrayLocked) {
                            return null;
                        }
                        return (
                            <>
                                {!isEPRReadOnly && (
                                    <button className={styles.table_btn} type="button" onClick={() => handleSplitClick(rowData)} title="Split">
                                        <img src={SplitIcon} />
                                    </button>
                                )}
                            </>
                        );
                    }}
                    />
                    <Column header="Action" body={renderAction} />
                </DataTable>
            </div>

            <ul className={styles.table_note_list}>
                <li className={styles.isplitrow}>
                    *Record not saved. Add required fields to update.
                </li>
                <li>
                    **Unassigned hours for the shift ({startTime} to {endTime}) : {performanceSummary.shiftDuration}
                </li>
            </ul>

            <Modal isOpen={showSplitPopup} isBlocking={false}>
                <div className={styles.modalPopup_header}>
                    <h2>Split Time</h2>
                </div>
                <div className={styles.modalPopup}>
                    <div className={styles.note_box}>
                        {splitRow && (
                            <>
                                <p>Split time must be between{' '} <b>{splitRow.Start}</b> {' '}and{' '} <b>{splitRow.End}</b>.</p>
                                <input type="text" className={styles.popup_input_box} list="splitTimeList" value={splitTime} onChange={(e) => setSplitTime(e.target.value)} placeholder="Select Split Time" />
                                <datalist id="splitTimeList">
                                    {timeOptions.map((time) => (<option key={time} value={time} />))}
                                </datalist>
                            </>
                        )}
                    </div>
                </div>
                <div className={styles.popup_button_box}>
                    <button type="button" className={styles.popup_button_1} onClick={confirmSplit}>
                        OK
                    </button>
                    <button type="button" className={styles.popup_button_2} onClick={() => { setShowSplitPopup(false); setSplitRow(null); setSplitTime(''); }}>
                        Cancel
                    </button>
                </div>
            </Modal>
        </>
    );
};

export default PerformanceTable;