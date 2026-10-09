import * as React from 'react';
import styles from '../css/Common.module.scss';
import SharePointService from '../service/SharePointService';
import { Modal } from '@fluentui/react';
import { SPFI } from '@pnp/sp';

interface IProps {
    stations: any[]; SelectedStations: any; size: any[]; grade: any[]; spService: SharePointService; leadsMember: any[]; sp: SPFI; hasPerformanceData?: boolean;
    performanceRows?: any[]; shiftHours: any; importantNotes: any; EPRAlertsMessages: any;
    onEPRLoaded?: (eprId: number, station: string, date: string, shift: string, startTime: string, endTime: string) => Promise<void> | void;
    onSubmitForApproval?: (eprId: number, station: string) => Promise<boolean>;
    onReadOnlyChange?: (readOnly: boolean) => void;
    onLoadingChange?: (loading: boolean) => void;
    onTimeChange?: (startTime: string, endTime: string) => void;
}

const EPRHeader: React.FC<IProps> = ({
    stations, SelectedStations, size, grade, spService, leadsMember, sp, hasPerformanceData, performanceRows, shiftHours, importantNotes, EPRAlertsMessages,
    onEPRLoaded, onSubmitForApproval, onReadOnlyChange, onLoadingChange, onTimeChange }) => {

    const [isEPRStarted, setIsEPRStarted] = React.useState<boolean>(false);
    const [selectedStation, setSelectedStation] = React.useState<string>(SelectedStations || (stations.length > 0 ? stations[0].Title : ''));
    const [selectedDate, setSelectedDate] = React.useState<string>(() => {
        const today = new Date();
        const year = today.getFullYear();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const day = String(today.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    });
    const [selectedShift, setSelectedShift] = React.useState<string>('Shift 1');
    const [currentStartTime, setCurrentStartTime] = React.useState<string>('');
    const [currentEndTime, setCurrentEndTime] = React.useState<string>('');
    const [currentEPR, setCurrentEPR] = React.useState<any | null>(null);
    const [operatorName, setOperatorName] = React.useState<string>('');
    const [selectedSize, setSelectedSize] = React.useState<string>('');
    const [selectedGrade, setSelectedGrade] = React.useState<string>('');
    const [operatorSignature, setOperatorSignature] = React.useState<string>('');
    const [levelOneApprovalStatus, setLevelOneApprovalStatus] = React.useState<string>('');
    // const [levelOneApprover, setLevelOneApprover] = React.useState<any | null>(null);
    const [levelOneRemark, setLevelOneRemark] = React.useState<string>('');
    const [isEPRReadOnly, setIsEPRReadOnly] = React.useState<boolean>(false);
    const [isClosed, setIsClosed] = React.useState<boolean>(false);
    const [isCloseShiftPending, setIsCloseShiftPending] = React.useState<boolean>(false);
    const [isSubmitForApproval, setIsSubmitForApproval] = React.useState<boolean>(false);

    const [startTimeError, setStartTimeError] = React.useState<string>('');
    const [endTimeError, setEndTimeError] = React.useState<string>('');

    const [isCurrentUserApprover, setIsCurrentUserApprover] = React.useState<boolean>(false);
    const [approvalAction, setApprovalAction] = React.useState<'Approved' | 'Rejected' | null>(null);
    const [showApprovalPopup, setShowApprovalPopup] = React.useState<boolean>(false);

    const [isSaving, setIsSaving] = React.useState<boolean>(false);
    const [showEPRPopup, setShowEPRPopup] = React.useState<boolean>(false);
    const [previousShiftEndTime, setPreviousShiftEndTime] = React.useState<string>('');
    const [eprPopupType, setEprPopupType] = React.useState<'create' | 'gap' | 'overlap' | 'runStop' | 'updateRunStop' | 'updateTimeWarning' | 'submitRunStop' | 'timeError' | 'dataTimeError' | 'closeShift'>('create');
    const [eprPopupMinutes, setEprPopupMinutes] = React.useState<number>(0);
    const [timeUpdateShiftWarning, setTimeUpdateShiftWarning] = React.useState<'gap' | 'overlap' | null>(null);
    const [runStopBeforeStartLogs, setRunStopBeforeStartLogs] = React.useState<any[]>([]);
    const [runStopAfterEndLogs, setRunStopAfterEndLogs] = React.useState<any[]>([]);
    const [showSuccessPopup, setShowSuccessPopup] = React.useState<boolean>(false);
    const [currentTime, setCurrentTime] = React.useState<Date>(new Date());
    const urlEPRLoadedRef = React.useRef<boolean>(false);

    /* Start of Utility Functions */
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

    const normalizeTime = (value: string): string => {
        if (!value) {
            return '';
        }
        const cleaned = value.trim().replace(/\s+/g, ' ').toUpperCase();
        const match = cleaned.match(/^(0?[1-9]|1[0-2]):([0-5][0-9])\s*(AM|PM)$/);
        if (!match) {
            return cleaned;
        }
        const hour = parseInt(match[1], 10);
        const minute = match[2];
        const period = match[3];
        return `${hour}:${minute} ${period}`;
    };

    const toMinutes = (value: string): number => {
        const normalized = normalizeTime(value);
        const parts = normalized.split(' ');
        if (parts.length !== 2) {
            return -1;
        }
        const time = parts[0];
        const period = parts[1];
        let [hour, minute] = time.split(':').map(Number);
        if (period === 'AM' && hour === 12) {
            hour = 0;
        }
        if (period === 'PM' && hour !== 12) {
            hour += 12;
        }

        return (hour * 60) + minute;
    };

    const getShiftConfig = (shift: string): any | null => {
        return (shiftHours.find((item: any) => item.ShiftType === shift) || null);
    };

    const getDefaultStartTime = (shift: string): string => {
        const config = getShiftConfig(shift);
        return config?.StartTime ? normalizeTime(config.StartTime) : '';
    };

    const getDefaultEndTime = (shift: string): string => {
        const config = getShiftConfig(shift);
        return config?.EndTime ? normalizeTime(config.EndTime) : '';
    };

    const getAlertText = (placement: string, values: Record<string, string | number> = {}): string => {
        const alert = (EPRAlertsMessages || []).find((item: any) => String(item.Placement || '').trim().toLowerCase() === placement.trim().toLowerCase());
        if (!alert?.Title) {
            console.warn(`EPR alert message not found for placement: ${placement}`);
            return '';
        }

        let message = String(alert.Title);
        Object.entries(values).forEach(
            ([key, value]) => { message = message.replace(new RegExp(`\\{${key}\\}`, 'g'), String(value)); }
        );

        return message;
    };


    const renderAlertMessage = (placement: string, values: Record<string, string | number> = {}): React.ReactNode => {
        const message = getAlertText(placement, values);
        if (!message) {
            return null;
        }
        const parts = message.split(/(\{[^}]+\})/g);
        return (
            <>
                {parts.map(
                    (part: string, index: number) => {
                        const match = part.match(/^\{(.+)\}$/);
                        if (match) {
                            const key = match[1];
                            return (
                                <b key={index}>
                                    {values[key] !== undefined ? values[key] : part}
                                </b>
                            );
                        }

                        return (
                            <React.Fragment key={index}>
                                {part}
                            </React.Fragment>
                        );
                    }
                )}
            </>
        );
    };


    /* End of Utility Functions */

    /* Start of UseEffect Hooks */
    React.useEffect(() => {
        const timer = window.setInterval(() => {
            setCurrentTime(new Date());
        }, 30000);

        return () => {
            window.clearInterval(timer);
        };
    }, []);

    React.useEffect(() => {
        const eprId = getEPRIdFromUrl();
        if (eprId) {
            if (urlEPRLoadedRef.current) {
                return;
            }
            urlEPRLoadedRef.current = true;
            void loadEPRFromUrl(eprId);
            return;
        }

        if (!selectedStation) { return; }
        void checkEPR(selectedStation, selectedDate, selectedShift);

    }, [selectedDate, selectedShift, shiftHours]);
    /* End of UseEffect Hooks */

    const isApprovalDisabled = (): boolean => {
        if (!selectedDate || !currentStartTime || !currentEndTime) {
            return true;
        }

        const startMinutes = toMinutes(currentStartTime);
        const endMinutes = toMinutes(currentEndTime);

        if (startMinutes < 0 || endMinutes < 0) {
            return true;
        }

        const [year, month, day] = selectedDate.split('-').map(Number);

        const endDateTime = new Date(
            year,
            month - 1,
            day,
            Math.floor(endMinutes / 60),
            endMinutes % 60,
            0
        );

        // Overnight shift
        if (endMinutes <= startMinutes) {
            endDateTime.setDate(endDateTime.getDate() + 1);
        }

        return currentTime.getTime() < endDateTime.getTime();
    };

    const getEPRIdFromUrl = (): number | null => {
        const query = window.location.search;
        if (!query) {
            return null;
        }

        const value = query.substring(1).trim();
        if (!value) {
            return null;
        }

        const id = Number(value);
        return Number.isFinite(id) ? id : null;
    };

    const SetUPAllData = async (epr: any, isStarted: boolean): Promise<void> => {
        let date = selectedDate;

        const submitForApproval = epr.SubmitForApproval === true;
        const closed = epr.IsClosed === true;
        const readOnly = closed || submitForApproval;
        const shift = epr.Shifts || 'Shift 1';
        const startTime = epr.StartTime || getDefaultStartTime(shift);
        const endTime = epr.EndTime || getDefaultEndTime(shift);
        const Id = epr.Id || 0;

        // Get the actual EPR date first.
        if (epr.Date) {
            const eprDate = new Date(epr.Date);

            if (!isNaN(eprDate.getTime())) {
                const year = eprDate.getFullYear();
                const month = String(eprDate.getMonth() + 1).padStart(2, '0');
                const day = String(eprDate.getDate()).padStart(2, '0');

                date = `${year}-${month}-${day}`;
            }
        }

        // Shift 1 has no previous-shift dependency.
        // Shift 2 must use Shift 1 from the SAME EPR date.
        if (shift === 'Shift 2') {
            const previousData = await spService.getEPRByDate(
                selectedStation,
                date,
                'Shift 1'
            );

            if (previousData && previousData.length > 0) {
                setPreviousShiftEndTime(
                    previousData[0].EndTime || ''
                );
            }
            else {
                setPreviousShiftEndTime('');
            }
        }
        else {
            setPreviousShiftEndTime('');
        }

        setSelectedDate(date);
        setSelectedShift(shift);
        setCurrentStartTime(startTime);
        setCurrentEndTime(endTime);
        setOperatorName(epr.InspectorName || '');
        setOperatorSignature(epr.LeadOperatorSignature || '');
        setSelectedSize(epr.Size || '');
        setSelectedGrade(epr.Grade || '');
        setLevelOneApprovalStatus(epr.LevelOneApprovalStatus || '');
        setLevelOneRemark(epr.LevelOneRemark || '');
        setIsSubmitForApproval(submitForApproval);
        setIsClosed(closed);
        setCurrentEPR(epr);
        setIsEPRStarted(isStarted);
        setIsEPRReadOnly(readOnly);

        onReadOnlyChange?.(readOnly);

        if (isStarted) {
            await onEPRLoaded?.(
                Id,
                selectedStation,
                date,
                shift,
                startTime,
                endTime
            );
        }

        if (readOnly) {
            await checkCurrentUserApprover(selectedStation);
        }
    };

    const loadEPRFromUrl = async (eprId: number): Promise<void> => {
        try {
            const existing = await spService.getEPRById(SelectedStations, eprId);
            if (!existing) {
                console.error('EPR not found:', eprId);
                return;
            }
            setIsEPRStarted(true);
            SetUPAllData(existing, true);
        }
        catch (error) {
            console.error('Error loading EPR from URL:', error);
        }
    };

    const checkEPR = async (station: string, date: string, shift: string): Promise<void> => {
        if (!station || !date || !shift) {
            return;
        }
        try {
            const currentData = await spService.getEPRByDate(station, date, shift);
            if (currentData && currentData.length > 0) {
                SetUPAllData(currentData[0], false);
                return;
            }
            else {
                setCurrentEPR(null);
                setIsClosed(false);
                setIsSubmitForApproval(false);
                setIsEPRReadOnly(false);
                onReadOnlyChange?.(false);

                setCurrentEndTime(getDefaultEndTime(shift));

                if (shift === 'Shift 1') {
                    setPreviousShiftEndTime('');
                    setCurrentStartTime(getDefaultStartTime('Shift 1'));
                }
                else {
                    const previousShift = 'Shift 1';
                    const previousData = await spService.getEPRByDate(station, date, previousShift);
                    if (previousData && previousData.length > 0) {
                        const previousEPR = previousData[0];
                        const previousEndTime = previousEPR.EndTime || getDefaultStartTime('Shift 2');
                        setPreviousShiftEndTime(previousEndTime);
                        setCurrentStartTime(previousEndTime);
                    }
                    else {
                        setPreviousShiftEndTime('');
                        setCurrentStartTime(getDefaultStartTime('Shift 2'));
                    }
                }
            }

        }
        catch (error) {
            console.error('Error checking EPR:', error);
            setCurrentEPR(null);
            setIsClosed(false);
            setIsSubmitForApproval(false);
            setIsCloseShiftPending(false);
            setIsSubmitForApproval(false);
            setIsEPRReadOnly(false);
            onReadOnlyChange?.(false);
        }
    };



    const validateShiftTimes = (): boolean => {
        const normalizedStart = normalizeTime(currentStartTime);
        const normalizedEnd = normalizeTime(currentEndTime);

        if (!normalizedStart) {
            setStartTimeError(getAlertText('EPR_START_REQUIRED'));
            setEprPopupType('timeError');
            setShowEPRPopup(true);
            return false;
        }

        if (!normalizedEnd) {
            setEndTimeError(getAlertText('EPR_END_REQUIRED'));
            setEprPopupType('timeError');
            setShowEPRPopup(true);
            return false;
        }

        const startMinutes = toMinutes(normalizedStart);
        const endMinutes = toMinutes(normalizedEnd);

        if (startMinutes < 0) {
            setStartTimeError(getAlertText('EPR_INVALID_START_TIME'));
            setEprPopupType('timeError');
            setShowEPRPopup(true);
            return false;
        }

        if (endMinutes < 0) {
            setEndTimeError(getAlertText('EPR_INVALID_END_TIME'));
            setEprPopupType('timeError');
            setShowEPRPopup(true);
            return false;
        }

        const shiftConfig = getShiftConfig(selectedShift);
        const configuredStart = shiftConfig?.StartTime ? toMinutes(normalizeTime(shiftConfig.StartTime)) : -1;
        const configuredEnd = shiftConfig?.EndTime ? toMinutes(normalizeTime(shiftConfig.EndTime)) : -1;
        const isOvernightShift = configuredStart >= 0 && configuredEnd >= 0 && configuredEnd <= configuredStart;

        if (!isOvernightShift && endMinutes < startMinutes) {
            setEndTimeError(getAlertText('EPR_END_BEFORE_START'));
            setEprPopupType('timeError');
            setShowEPRPopup(true);
            return false;
        }

        if (startMinutes === endMinutes) {
            setEndTimeError(getAlertText('EPR_START_END_SAME'));
            setEprPopupType('timeError');
            setShowEPRPopup(true);
            return false;
        }

        setStartTimeError('');
        setEndTimeError('');
        return true;
    };

    const handleStationChange = (station: string): void => {
        const selectedStationItem = stations.find((item: any) => item.Title === station);
        if (!selectedStationItem) {
            return;
        }
        setSelectedStation(station);
        if (selectedStationItem.URL) {
            window.location.href = selectedStationItem.URL.Url;
        }
    };

    const handleDateChange = (date: string): void => {
        setSelectedDate(date);
        setStartTimeError('');
        setEndTimeError('');
    };

    const handleShiftChange = (shift: string): void => {
        setSelectedShift(shift);
        setStartTimeError('');
        setEndTimeError('');
    };

    const validatePreviousShiftTime = (): { isValid: boolean; type: 'gap' | 'overlap' | null; minutes: number; } => {
        if (!previousShiftEndTime) {
            return { isValid: true, type: null, minutes: 0 };
        }
        const previousEnd = toMinutes(previousShiftEndTime);
        const currentStart = toMinutes(currentStartTime);

        if (previousEnd < 0 || currentStart < 0) {
            return { isValid: true, type: null, minutes: 0 };
        }
        const difference = currentStart - previousEnd;

        if (difference === 0) {
            return { isValid: true, type: null, minutes: 0 };
        }

        if (difference > 0) {
            return { isValid: false, type: 'gap', minutes: difference };
        }

        return { isValid: false, type: 'overlap', minutes: Math.abs(difference) };
    };

    const parseRunStopDateTime = (value: string): Date | null => {
        if (!value) {
            return null;
        }

        const [datePart, timePart] = value.trim().split(' ');

        if (!datePart || !timePart) {
            return null;
        }

        const [month, day, year] = datePart.split('/').map(Number);
        const [hours, minutes, seconds] = timePart.split(':').map(Number);

        if (!month || !day || !year || isNaN(hours) || isNaN(minutes) || isNaN(seconds)) {
            return null;
        }
        return new Date(year, month - 1, day, hours, minutes, seconds);
    };

    const createDateTime = (date: string, time: string): Date | null => {
        if (!date || !time) {
            return null;
        }

        const [year, month, day] = date.split('-').map(Number);
        const totalMinutes = toMinutes(time);

        if (totalMinutes < 0) {
            return null;
        }

        const hours = Math.floor(totalMinutes / 60);
        const minutes = totalMinutes % 60;
        return new Date(year, month - 1, day, hours, minutes, 0);
    };

    const getShiftBeforeStartWindow = (shift: string): string => {
        const config = getShiftConfig(shift);
        return config?.BeforeStartWindowTime ? normalizeTime(config.BeforeStartWindowTime) : '';
    };

    const getShiftAfterEndWindow = (shift: string): string => {
        const config = getShiftConfig(shift);
        return config?.AfterEndWindowTime ? normalizeTime(config.AfterEndWindowTime) : '';
    };

    const getRunStopFieldMapping = (station: string): { timestampField: string; dateField: string; machineStateField: string; } => {
        switch (station) {
            case 'HeatTreat': case 'ERW Mill': return { timestampField: 'field_4', dateField: 'field_10', machineStateField: 'field_5' };
            case 'Hydro': case 'NDT': case 'Threadline1': case 'Threadline2': return { timestampField: 'field_1', dateField: 'field_3', machineStateField: 'field_4' };
            default: return { timestampField: 'field_4', dateField: 'field_10', machineStateField: 'field_5' };
        }
    };

    const validateRunStopLogs = (runStopLogs: any[]): { hasBeforeStart: boolean; hasAfterEnd: boolean; beforeStartLogs: any[]; afterEndLogs: any[]; } => {
        if (!runStopLogs || runStopLogs.length === 0) {
            return { hasBeforeStart: false, hasAfterEnd: false, beforeStartLogs: [], afterEndLogs: [] };
        }

        const startDateTime = createDateTime(selectedDate, currentStartTime);
        const calculatedEndDateTime = createDateTime(selectedDate, currentEndTime);

        if (startDateTime === null || calculatedEndDateTime === null) {
            return { hasBeforeStart: false, hasAfterEnd: false, beforeStartLogs: [], afterEndLogs: [] };
        }

        const endDateTime = new Date(calculatedEndDateTime.getTime());
        if (endDateTime.getTime() <= startDateTime.getTime()) { endDateTime.setDate(endDateTime.getDate() + 1); }

        let beforeStartWindow: Date;
        let afterEndWindow: Date;

        const beforeStartTime = getShiftBeforeStartWindow(selectedShift);
        const afterEndTime = getShiftAfterEndWindow(selectedShift);

        if (!beforeStartTime || !afterEndTime) {
            return { hasBeforeStart: false, hasAfterEnd: false, beforeStartLogs: [], afterEndLogs: [] };
        }

        beforeStartWindow = createDateTime(selectedDate, beforeStartTime)!;
        afterEndWindow = createDateTime(selectedDate, afterEndTime)!;

        if (beforeStartWindow.getTime() >= startDateTime.getTime()) {
            beforeStartWindow.setDate(beforeStartWindow.getDate() - 1);
        }

        if (afterEndWindow.getTime() < endDateTime.getTime()) {
            afterEndWindow.setDate(afterEndWindow.getDate() + 1);
        }

        const beforeStartLogs: any[] = [];
        const afterEndLogs: any[] = [];

        runStopLogs.forEach((log: any) => {
            const {
                timestampField
            } = getRunStopFieldMapping(selectedStation);

            const dateField = log[timestampField];
            if (!dateField) {
                return;
            }
            const logDateTime = parseRunStopDateTime(dateField);

            if (logDateTime === null) {
                return;
            }

            const logTime = logDateTime.getTime();

            if (logTime >= beforeStartWindow.getTime() && logTime < startDateTime.getTime()) {
                beforeStartLogs.push(log);
                return;
            }

            if (logTime > endDateTime.getTime() && logTime <= afterEndWindow.getTime()) {
                afterEndLogs.push(log);
            }
        });

        return { hasBeforeStart: beforeStartLogs.length > 0, hasAfterEnd: afterEndLogs.length > 0, beforeStartLogs, afterEndLogs };
    };

    const getRunStopLogs = async (): Promise<any[]> => {
        try {
            const selectedDateLogs = await spService.getEPRrunStopLogs(selectedStation, selectedDate);
            let allLogs = selectedDateLogs || [];

            if (selectedShift === 'Shift 2') {
                const nextDate = new Date(`${selectedDate}T00:00:00`);
                nextDate.setDate(nextDate.getDate() + 1);
                const year = nextDate.getFullYear();
                const month = String(nextDate.getMonth() + 1).padStart(2, '0');
                const day = String(nextDate.getDate()).padStart(2, '0');
                const nextDateString = `${year}-${month}-${day}`;
                const nextDateLogs = await spService.getEPRrunStopLogs(selectedStation, nextDateString);
                allLogs = [...allLogs, ...(nextDateLogs || [])];
            }

            const uniqueLogs = Array.from(new Map(allLogs.map((log: any) => [log.Id, log])).values());
            return uniqueLogs;
        }
        catch (error) {
            console.error('Error loading Run/Stop logs:', error);
            return [];
        }
    };

    const handleGo = async (): Promise<void> => {
        if (currentEPR) {
            const eprStartTime = currentEPR.StartTime || getDefaultStartTime(selectedShift);
            const eprEndTime = currentEPR.EndTime || getDefaultEndTime(selectedShift);
            setCurrentStartTime(eprStartTime);
            setCurrentEndTime(eprEndTime);
            if (onEPRLoaded) {
                await onEPRLoaded(currentEPR.Id, selectedStation, selectedDate, selectedShift, eprStartTime, eprEndTime);
            }
            setIsEPRStarted(true);
            return;
        }
        if (!validateShiftTimes()) {
            return;
        }
        if (selectedShift !== 'Shift 1') {
            const previousShiftValidation = validatePreviousShiftTime();

            if (!previousShiftValidation.isValid) {
                setEprPopupType(previousShiftValidation.type!);
                setEprPopupMinutes(previousShiftValidation.minutes);
                setShowEPRPopup(true);
                return;
            }
        }

        const runStopLogs = await getRunStopLogs();
        const runStopValidation = validateRunStopLogs(runStopLogs);

        if (runStopValidation.hasBeforeStart || runStopValidation.hasAfterEnd) {
            setRunStopBeforeStartLogs(runStopValidation.beforeStartLogs);
            setRunStopAfterEndLogs(runStopValidation.afterEndLogs);
            setEprPopupType('runStop');
            setShowEPRPopup(true);
            return;
        }

        setEprPopupType('create');
        setShowEPRPopup(true);
    };

    const handleEPRPopupCancel = (): void => {
        setShowEPRPopup(false);
        setIsSaving(false);
    };

    const handleEPRPopupOK = async (): Promise<void> => {
        if (eprPopupType === 'timeError' || eprPopupType === 'dataTimeError') {
            setShowEPRPopup(false);
            return;
        }

        if ((eprPopupType === 'gap' || eprPopupType === 'overlap') && currentEPR?.Id) {
            const runStopLogs = await getRunStopLogs();
            const runStopValidation = validateRunStopLogs(runStopLogs);
            if (runStopValidation.hasBeforeStart || runStopValidation.hasAfterEnd) {
                setRunStopBeforeStartLogs(runStopValidation.beforeStartLogs);
                setRunStopAfterEndLogs(runStopValidation.afterEndLogs);
                setEprPopupType('updateRunStop');
                return;
            }
            setShowEPRPopup(false);
            await saveUpdatedEPRTime(normalizeTime(currentStartTime), normalizeTime(currentEndTime));
            return;
        }

        if (eprPopupType === 'updateTimeWarning') {
            setShowEPRPopup(false); await saveUpdatedEPRTime(normalizeTime(currentStartTime), normalizeTime(currentEndTime));
            return;
        }

        if (eprPopupType === 'updateRunStop') {
            setShowEPRPopup(false);
            await saveUpdatedEPRTime(normalizeTime(currentStartTime), normalizeTime(currentEndTime));
            return;
        }

        if (eprPopupType === 'closeShift') {
            setShowEPRPopup(false);
            await saveEPRWithClose();
            return;
        }

        if (eprPopupType === 'submitRunStop') {
            setShowEPRPopup(false);
            await submitEPRForApproval();
            return;
        }

        setShowEPRPopup(false);
        await createNewEPR();
        setIsSaving(false);
    };

    const createNewEPR = async (): Promise<void> => {
        try {

            const selectedStationItem = stations.find((item: any) => item.Title === selectedStation);
            const eprPayload = {
                Title: `${selectedStation}-${selectedDate}-${selectedShift}`,
                StationName: selectedStation,
                StationId: selectedStationItem?.Id || null,
                Date: selectedDate,
                StartTime: normalizeTime(currentStartTime),
                EndTime: normalizeTime(currentEndTime),
                Shifts: selectedShift
            };

            const createdEPR = await spService.createEPR(selectedStation, eprPayload);
            setCurrentEPR(createdEPR);
            setIsEPRStarted(true);
            await onEPRLoaded?.(createdEPR.Id, selectedStation, selectedDate, selectedShift, currentStartTime, currentEndTime);
        }
        catch (error) {
            console.error('Error creating EPR:', error);
        }
    };

    const saveUpdatedEPRTime = async (startTime: string, endTime: string): Promise<void> => {
        if (!currentEPR?.Id) {
            return;
        }

        try {
            setIsSaving(true);
            await spService.updateEPR(selectedStation, currentEPR.Id, { StartTime: startTime, EndTime: endTime });
            setCurrentEPR({ ...currentEPR, StartTime: startTime, EndTime: endTime });
            setCurrentStartTime(startTime);
            setCurrentEndTime(endTime);
            onTimeChange?.(startTime, endTime);
            setShowSuccessPopup(true);
        }
        catch (error) {
            console.error('Error updating EPR time:', error);
            setCurrentStartTime(currentEPR?.StartTime || '');
            setCurrentEndTime(currentEPR?.EndTime || '');
            setEprPopupType('timeError');
            setStartTimeError(getAlertText('EPR_UPDATE_TIME_ERROR'));
            setShowEPRPopup(true);
        }
        finally {
            setIsSaving(false);
        }
    };

    const handleTimeUpdate = async (type: 'start' | 'end'): Promise<void> => {
        const newStartTime = normalizeTime(currentStartTime);
        const newEndTime = normalizeTime(currentEndTime);

        if (newStartTime === normalizeTime(currentEPR?.StartTime || '') && newEndTime === normalizeTime(currentEPR?.EndTime || '')) {
            return;
        }

        const currentEPRStart = normalizeTime(currentEPR?.StartTime || '');
        const currentEPREnd = normalizeTime(currentEPR?.EndTime || '');

        const hasDataAtStartBoundary = (performanceRows || []).some((row: any) => (row.byHoursId || row.IsSplit === 'Yes') && normalizeTime(row.Start) === currentEPRStart);
        const hasDataAtEndBoundary = (performanceRows || []).some((row: any) => (row.byHoursId || row.IsSplit === 'Yes') && normalizeTime(row.End) === currentEPREnd);

        if (type === 'start' && hasDataAtStartBoundary) {
            setCurrentStartTime(currentEPR?.StartTime || currentStartTime);
            setEprPopupType('dataTimeError');
            setShowEPRPopup(true);
            return;
        }

        if (type === 'end' && hasDataAtEndBoundary) {
            setCurrentEndTime(currentEPR?.EndTime || currentEndTime);
            setEprPopupType('dataTimeError');
            setShowEPRPopup(true);
            return;
        }

        const startMinutes = toMinutes(newStartTime);
        const endMinutes = toMinutes(newEndTime);

        if (startMinutes < 0) {
            setCurrentStartTime(currentEPR?.StartTime || '');
            setStartTimeError(getAlertText('EPR_INVALID_START_TIME'));
            setEprPopupType('timeError');
            setShowEPRPopup(true);
            return;
        }

        if (endMinutes < 0) {
            setCurrentEndTime(currentEPR?.EndTime || '');
            setEndTimeError(getAlertText('EPR_INVALID_END_TIME'));
            setEprPopupType('timeError');
            setShowEPRPopup(true);
            return;
        }

        if (selectedShift === 'Shift 1' && endMinutes < startMinutes) {
            if (type === 'start') {
                setCurrentStartTime(currentEPR?.StartTime || '');
            }
            else {
                setCurrentEndTime(currentEPR?.EndTime || '');
            }
            setEndTimeError(getAlertText('EPR_END_BEFORE_START'));
            setEprPopupType('timeError');
            setShowEPRPopup(true);
            return;
        }

        if (startMinutes === endMinutes) {
            if (type === 'start') {
                setCurrentStartTime(currentEPR?.StartTime || '');
            }
            else {
                setCurrentEndTime(currentEPR?.EndTime || '');
            }
            setEndTimeError(getAlertText('EPR_START_END_SAME'));
            setEprPopupType('timeError');
            setShowEPRPopup(true);
            return;
        }

        let shiftWarning: 'gap' | 'overlap' | null = null;
        let shiftWarningMinutes = 0;

        if (type === 'start' && selectedShift === 'Shift 2') {
            const previousShiftValidation = validatePreviousShiftTime();
            if (!previousShiftValidation.isValid) {
                shiftWarning = previousShiftValidation.type;
                shiftWarningMinutes = previousShiftValidation.minutes;
            }
        }

        const runStopLogs = await getRunStopLogs();
        const runStopValidation = validateRunStopLogs(runStopLogs);
        const hasRunStopWarning = runStopValidation.hasBeforeStart || runStopValidation.hasAfterEnd;

        if (shiftWarning || hasRunStopWarning) {
            setTimeUpdateShiftWarning(shiftWarning);
            setEprPopupMinutes(shiftWarningMinutes);
            setRunStopBeforeStartLogs(runStopValidation.beforeStartLogs);
            setRunStopAfterEndLogs(runStopValidation.afterEndLogs);
            setEprPopupType('updateTimeWarning');
            setShowEPRPopup(true);
            return;
        }

        if (!currentEPR?.Id) { return; }
        await saveUpdatedEPRTime(newStartTime, newEndTime);
    };

    const updateFinalEPRData = async (
        isClosed: boolean,
        submitForApproval: boolean
    ): Promise<void> => {

        if (!currentEPR?.Id) {
            return;
        }

        const data = {
            Title: currentEPR.Title || `${selectedStation}-${selectedDate}-${selectedShift}`,

            StationName: selectedStation,

            StationId:
                currentEPR.StationId ||
                stations.find((item: any) => item.Title === selectedStation)?.Id ||
                null,

            Date: selectedDate,

            StartTime: normalizeTime(currentStartTime),

            EndTime: normalizeTime(currentEndTime),

            Shifts: selectedShift,

            InspectorName: operatorName,

            Size: selectedSize,

            Grade: selectedGrade,

            LeadOperatorSignature: operatorSignature,

            IsClosed: isClosed,

            SubmitForApproval: submitForApproval,

            LevelOneApprovalStatus: levelOneApprovalStatus || '',

            LevelOneRemark: levelOneRemark,

            LevelOneApproverId:
                currentEPR.LevelOneApproverId || null
        };

        await spService.updateEPR(
            selectedStation,
            currentEPR.Id,
            data
        );

        setCurrentEPR({
            ...currentEPR,
            ...data
        });
    };

    const saveEPRWithClose = async (): Promise<void> => {
        if (!currentEPR?.Id) {
            return;
        }

        try {
            setIsSaving(true);

            await updateFinalEPRData(
                true,
                false
            );

            setIsClosed(true);
            setIsCloseShiftPending(false);
            setIsEPRReadOnly(true);

            onReadOnlyChange?.(true);

            setShowSuccessPopup(true);
        }
        catch (error) {
            console.error(
                'Error closing EPR:',
                error
            );
        }
        finally {
            setIsSaving(false);
        }
    };

    const handleSave = async (): Promise<void> => {
        if (isCloseShiftPending && !isClosed) {
            setEprPopupType('closeShift');
            setShowEPRPopup(true);
            return;
        }

        if (!validateShiftTimes()) {
            return;
        }

        if (!currentEPR?.Id) {
            console.error('Current EPR ID is missing.');
            return;
        }

        try {
            setIsSaving(true);
            const data = { InspectorName: operatorName, Size: selectedSize, Grade: selectedGrade, StartTime: normalizeTime(currentStartTime), EndTime: normalizeTime(currentEndTime), IsClosed: isClosed, LeadOperatorSignature: operatorSignature };
            await spService.updateEPR(selectedStation, currentEPR.Id, data);
            setCurrentEPR({ ...currentEPR, ...data });
            setShowSuccessPopup(true);
            setIsEPRReadOnly(isClosed);
            onReadOnlyChange?.(isClosed);
        }
        catch (error) {
            console.error('Error updating EPR:', error);
        }
        finally {

            setIsSaving(false);

        }
    };

    const handleCancel = (): void => {
        location.reload();
    };

    const submitEPRForApproval = async (): Promise<void> => {

        const eprId = currentEPR?.Id || currentEPR?.ID;

        if (!eprId) {
            console.error(
                'EPR ID not found:',
                currentEPR
            );

            alert(
                getAlertText('EPR_ID_MISSING')
            );

            return;
        }

        try {
            setIsSaving(true);

            const submitted =
                await onSubmitForApproval?.(
                    eprId,
                    selectedStation
                );

            if (submitted) {

                await updateFinalEPRData(
                    true,
                    true
                );

                setIsSubmitForApproval(true);

                setIsClosed(true);

                setIsCloseShiftPending(false);

                setIsEPRReadOnly(true);

                onReadOnlyChange?.(true);

                await checkCurrentUserApprover(
                    selectedStation
                );
            }
        }
        catch (error) {
            console.error(
                'Error submitting EPR for approval:',
                error
            );
        }
        finally {
            setIsSaving(false);
        }
    };

    const handleSubmitForApproval = async (): Promise<void> => {
        const eprId = currentEPR?.Id || currentEPR?.ID;
        if (!eprId) {
            console.error('EPR ID not found:', currentEPR);
            alert(getAlertText('EPR_ID_MISSING'));
            return;
        }
        try {
            setIsSaving(true);
            const runStopLogs = await getRunStopLogs();
            const runStopValidation = validateRunStopLogs(runStopLogs);

            if (runStopValidation.hasBeforeStart || runStopValidation.hasAfterEnd) {
                setRunStopBeforeStartLogs(runStopValidation.beforeStartLogs);
                setRunStopAfterEndLogs(runStopValidation.afterEndLogs);
                setEprPopupType('submitRunStop');
                setShowEPRPopup(true);
                return;
            }
            setIsSaving(false);
            await submitEPRForApproval();
        }
        catch (error) {
            console.error('Error validating Run/Stop logs before approval:', error);
            setIsSaving(false);
        }
    };

    const checkCurrentUserApprover = async (station: string): Promise<void> => {
        try {
            const user = await sp.web.currentUser();
            const currentUserEmail = (user?.Email || user?.Email || '').toLowerCase();
            if (!currentUserEmail) {
                setIsCurrentUserApprover(false);
                return;
            }

            const participant = (leadsMember || []).find((item: any) => item.Stations?.Title?.toLowerCase() === station?.toLowerCase());
            if (!participant?.ApproverName) {
                setIsCurrentUserApprover(false);
                return;
            }

            const isApprover = participant.ApproverName.some((approver: any) => (approver.EMail || '').toLowerCase() === currentUserEmail);
            setIsCurrentUserApprover(isApprover);

        }

        catch (error) {
            console.error('Error checking current user approver:', error);
            setIsCurrentUserApprover(false);
        }
    };

    const handleApprove = async (): Promise<void> => {
        setApprovalAction('Approved');
        await handleApprovalSubmit('Approved');
    };

    const handleReject = async (): Promise<void> => {
        setApprovalAction('Rejected');
        await handleApprovalSubmit('Rejected');
    };

    const handleApprovalSubmit = async (action: 'Approved' | 'Rejected'): Promise<void> => {
        if (!currentEPR?.Id) return;
        try {
            setIsSaving(true);
            const currentUserId = (window as any)?._spPageContextInfo?.userId;
            const updateData = { LevelOneApprovalStatus: action, LevelOneApproverId: currentUserId, LevelOneRemark: levelOneRemark, IsClosed: action === 'Rejected' ? false : true, SubmitForApproval: action === 'Rejected' ? false : true };
            await spService.updateEPR(selectedStation, currentEPR.Id, updateData);
            setLevelOneApprovalStatus(action);
            setCurrentEPR({ ...currentEPR, LevelOneApprovalStatus: action, LevelOneRemark: levelOneRemark, IsClosed: action === 'Rejected' ? false : true, SubmitForApproval: action === 'Rejected' ? false : true });
            if (action === 'Rejected') {
                setIsClosed(false);
                setIsSubmitForApproval(false);
                setIsCloseShiftPending(false);
                setIsEPRReadOnly(false);
                onReadOnlyChange?.(false);
            }
            setShowApprovalPopup(true);
        }
        catch (error) {
            console.error('Error processing approval:', error);
        }
        finally {
            setIsSaving(false);
        }
    };

    return (
        <div className={styles.epr_header_row}>

            <div className={styles.epr_header_title_row}>
                <div className={styles.epr_ht_icon_box}>

                </div>
                <div className={styles.epr_ht_title_box}>
                    <h3>Equipment Performance Report</h3>
                    <p>Track equipment efficiency and operational performance in real-time.</p>
                </div>
                <div className={styles.epr_header_button_box}>
                    {!isEPRStarted && (
                        <button type="button" className={styles.epr_ht_go_btn} onClick={handleGo} disabled={isSaving}>
                            {isSaving ? 'LOADING...' : 'GO'}
                        </button>
                    )}
                    {isEPRStarted && (<>
                        {!isSubmitForApproval && (
                            <div className={styles.eprActionButtons}>
                                {(!isEPRReadOnly) && (<>
                                    <button type="button" className={styles.epr_ht_save_btn} onClick={handleSave}>
                                        Save
                                    </button>
                                    <button type="button" className={styles.epr_ht_cancel_btn} onClick={handleCancel}>
                                        Cancel
                                    </button>
                                </>)}
                                <label className={`${styles.isclose_btn} ${isClosed || isCloseShiftPending ? styles.checked : ''}`}>
                                    <input type="checkbox" checked={isClosed || isCloseShiftPending} disabled={isSubmitForApproval || isClosed} onChange={(event) => { const checked = event.target.checked; setIsCloseShiftPending(checked); }} />
                                    <span>
                                        {isClosed ? 'Shift Closed' : 'Close shift'}
                                    </span>
                                </label>
                                <button type="button" className={styles.epr_ht_approval_btn} onClick={handleSubmitForApproval} disabled={isSaving || isApprovalDisabled()}>
                                    Submit for Approval
                                </button>
                            </div>
                        )}
                        {isSubmitForApproval && isCurrentUserApprover && levelOneApprovalStatus !== 'Approved' && (
                            <div className={styles.eprActionButtons}>
                                <button type="button" className={styles.epr_ht_approve_btn} onClick={handleApprove} disabled={isSaving}>
                                    Approve
                                </button>
                                <button type="button" className={styles.epr_ht_reject_btn} onClick={handleReject} disabled={isSaving}>
                                    Reject and Send back
                                </button>
                            </div>
                        )}
                        {levelOneApprovalStatus === 'Approved' && (
                            <div className={styles.eprActionButtons}>
                                <button type="button" className={styles.epr_ht_approve_btn} disabled>
                                    Approved
                                </button>
                            </div>
                        )}
                        {isSubmitForApproval && levelOneApprovalStatus !== 'Approved' && !isCurrentUserApprover && (
                            <div className={styles.eprActionButtons}>
                                <button type="button" className={styles.epr_ht_pending_btn}>
                                    Pending Approval
                                </button>
                            </div>
                        )}
                    </>)}
                </div>
            </div>

            <ul className={styles.epr_filter_list}>
                <li>
                    <label className={styles.epr_fl_label}>Equipment/Station Name</label>
                    <select disabled={isEPRStarted} className={styles.epr_fl_textfld + ' ' + styles.epr_fl_choicefld} id="ddlStation" name="Station" value={selectedStation} onChange={(event) => { handleStationChange(event.target.value); }}>
                        <option value="">Select Station</option>
                        {stations.map((item, index) => (
                            <option key={item.Id || index} value={item.Title}>
                                {item.Title}
                            </option>
                        ))}
                    </select>
                </li>
                <li className={styles.boxof2}>
                    <div className={styles.epr_fl_halfwidth}>
                        <label className={styles.epr_fl_label}>Date</label>
                        <input disabled={isEPRStarted} className={styles.epr_fl_textfld} max={new Date().toLocaleDateString('en-CA')} id="txtDate" name="Date" type="date" value={selectedDate} onChange={(event) => { handleDateChange(event.target.value); }} />
                    </div>
                    <div className={styles.epr_fl_halfwidth}>
                        <label className={styles.epr_fl_label}>Shift</label>
                        <select disabled={isEPRStarted} className={styles.epr_fl_textfld + ' ' + styles.epr_fl_choicefld} id="ddlShift" name="Shifts" value={selectedShift} onChange={(event) => { handleShiftChange(event.target.value); }}>
                            <option value="Shift 1">Shift 1</option>
                            <option value="Shift 2">Shift 2</option>
                        </select>
                    </div>
                </li>
                <li className={styles.boxof2}>
                    <div className={styles.epr_fl_halfwidth}>
                        <label className={styles.epr_fl_label}>Start Time</label>
                        <input disabled={isClosed || isSubmitForApproval} className={`${styles.epr_fl_textfld} ${startTimeError ? styles.epr_fl_error_field : ''}`} id="StartTime" name="StartTime" list="StartTimeList" value={currentStartTime} autoComplete="off"
                            onChange={(event) => {
                                setCurrentStartTime(event.target.value);
                                setStartTimeError('');
                            }}
                            onBlur={() => { void handleTimeUpdate('start'); }}
                        />
                        <datalist id="StartTimeList">
                            {timeOptions.map((time) => (<option key={time} value={time} />))}
                        </datalist>
                    </div>
                    <div className={styles.epr_fl_halfwidth}>
                        <label className={styles.epr_fl_label}>End Time</label>
                        <input disabled={isClosed || isSubmitForApproval} className={`${styles.epr_fl_textfld} ${endTimeError ? styles.epr_fl_error_field : ''}`} id="EndTime" name="EndTime" list="EndTimeList" value={currentEndTime} autoComplete="off"
                            onChange={(event) => {
                                setCurrentEndTime(event.target.value);
                                setEndTimeError('');
                            }}
                            onBlur={() => { void handleTimeUpdate('end'); }}
                        />
                        <datalist id="EndTimeList">
                            {timeOptions.map((time) => (<option key={time} value={time} />))}
                        </datalist>
                    </div>
                </li>
                {isEPRStarted && (
                    <>
                        <li>
                            <label className={styles.epr_fl_label}>Operator / Inspector Name</label>
                            <input disabled={isEPRReadOnly} type="text" placeholder='Operator / Inspector Name' className={styles.epr_fl_textfld} value={operatorName} onChange={(event) => { setOperatorName(event.target.value); }} />
                        </li>
                        <li>
                            <label className={styles.epr_fl_label}>Size</label>
                            <select disabled={isEPRReadOnly} className={`${styles.epr_fl_textfld} ${styles.epr_fl_choicefld}`} id="ddlSize" name="Size" value={selectedSize} onChange={(event) => { setSelectedSize(event.target.value); }}>
                                <option value="">Select Size</option>
                                {size.map((item, index) => (
                                    <option key={item.Id || index} value={item.Title}>{item.Title}</option>
                                ))}
                            </select>
                        </li>
                        <li>
                            <label className={styles.epr_fl_label}>Grade</label>
                            <select disabled={isEPRReadOnly} className={`${styles.epr_fl_textfld} ${styles.epr_fl_choicefld}`} id="ddlGrade" name="Grade" value={selectedGrade} onChange={(event) => { setSelectedGrade(event.target.value); }}>
                                <option value="">Select Grade</option>
                                {grade.map((item, index) => (
                                    <option key={item.Id || index} value={item.Title}>{item.Title}</option>
                                ))}
                            </select>
                        </li>
                        <li className={styles.fullwidth + ' ' + styles.justifycenter}>
                            <label className={styles.epr_fl_label}>Lead Operator's Signature</label>
                            <input disabled={isEPRReadOnly} type="text" placeholder="Lead Operator's Signature" className={styles.epr_fl_textfld} value={operatorSignature} onChange={(event) => { setOperatorSignature(event.target.value); }} />
                        </li>
                    </>
                )}
                {isEPRStarted &&
                    (isCurrentUserApprover || levelOneApprovalStatus === 'Approved' || levelOneApprovalStatus === 'Rejected') && (
                        <li className={styles.fullwidth}>
                            <label className={styles.epr_fl_label}>
                                {(levelOneApprovalStatus === 'Approved' || levelOneApprovalStatus === 'Rejected') ? <>Plant Lead have {levelOneApprovalStatus} it.</> : <>Approval comment</>}
                            </label>
                            <textarea className={styles.epr_fl_textfld + ' ' + styles.epr_fl_textarea} value={levelOneRemark} onChange={(event) => setLevelOneRemark(event.target.value)} disabled={!isCurrentUserApprover || levelOneApprovalStatus === 'Approved'} placeholder="Enter feedback" />
                        </li>
                    )}
            </ul>

            <div className={styles.epr_note_box}>
                <h3>Important Notes:</h3>
                <div className={styles.importantNotes} dangerouslySetInnerHTML={{ __html: importantNotes[0]?.Description || '' }} />
            </div>

            <Modal isOpen={showEPRPopup} isBlocking={false}>
                <div className={styles.modalPopup_header}>
                    <h2>Equipment Performance Report</h2>
                </div>
                <div className={styles.modalPopup}>
                    <div className={styles.note_box}>

                        {eprPopupType === 'create' && (
                            <p>
                                {renderAlertMessage('EPR_CREATE_CONFIRMATION')}
                            </p>
                        )}
                        {eprPopupType === 'gap' && (
                            <p>
                                {renderAlertMessage('EPR_GAP_WARNING', { PreviousShiftEndTime: previousShiftEndTime, CurrentStartTime: currentStartTime, Minutes: eprPopupMinutes })}
                            </p>
                        )}

                        {eprPopupType === 'overlap' && (
                            <p>
                                {renderAlertMessage('EPR_OVERLAP_WARNING', { PreviousShiftEndTime: previousShiftEndTime, CurrentStartTime: currentStartTime, Minutes: eprPopupMinutes })}
                            </p>
                        )}

                        {eprPopupType === 'runStop' && (
                            <>
                                <p>{renderAlertMessage('EPR_RUNSTOP_WARNING')}</p>
                                {runStopBeforeStartLogs.length > 0 && (
                                    <div className={styles.runStopWarningSection}>
                                        <h4>
                                            {getAlertText('EPR_RUNSTOP_BEFORE_START_TITLE')}
                                        </h4>
                                        <ul>
                                            {runStopBeforeStartLogs.slice(0, 5).map((log: any, index: number) => {
                                                const {
                                                    timestampField
                                                } = getRunStopFieldMapping(selectedStation);

                                                const dateField = log[timestampField];
                                                return (
                                                    <li key={log.Id || index}>
                                                        {dateField}
                                                    </li>
                                                );
                                            })}
                                        </ul>
                                        {runStopBeforeStartLogs.length > 5 && (
                                            <span>
                                                {renderAlertMessage('EPR_RUNSTOP_MORE', { Count: runStopBeforeStartLogs.length - 5 })}
                                            </span>
                                        )}
                                    </div>
                                )}

                                {runStopAfterEndLogs.length > 0 && (
                                    <div className={styles.runStopWarningSection}>
                                        <h4>
                                            {getAlertText('EPR_RUNSTOP_AFTER_END_TITLE')}
                                        </h4>
                                        <ul>
                                            {runStopAfterEndLogs.slice(0, 5).map((log: any, index: number) => {
                                                const {
                                                    timestampField
                                                } = getRunStopFieldMapping(selectedStation);

                                                const dateField = log[timestampField];
                                                return (
                                                    <li key={log.Id || index}>
                                                        {dateField}
                                                    </li>
                                                );
                                            })}
                                        </ul>
                                        {runStopAfterEndLogs.length > 5 && (
                                            <span>
                                                {renderAlertMessage('EPR_RUNSTOP_MORE', { Count: runStopAfterEndLogs.length - 5 })}
                                            </span>
                                        )}
                                    </div>
                                )}
                                <p>
                                    {renderAlertMessage('EPR_RUNSTOP_CONTINUE')}
                                </p>
                            </>
                        )}

                        {eprPopupType === 'submitRunStop' && (
                            <>
                                <p>
                                    {renderAlertMessage('EPR_RUNSTOP_WARNING')}
                                </p>

                                {runStopBeforeStartLogs.length > 0 && (
                                    <div className={styles.runStopWarningSection}>
                                        <h4>
                                            {getAlertText('EPR_RUNSTOP_BEFORE_START_TITLE')}
                                        </h4>
                                        <ul>
                                            {runStopBeforeStartLogs.slice(0, 5).map((log: any, index: number) => {
                                                const {
                                                    timestampField
                                                } = getRunStopFieldMapping(selectedStation);

                                                const dateField = log[timestampField];
                                                return (
                                                    <li key={log.Id || index}>
                                                        {dateField}
                                                    </li>
                                                );
                                            })}
                                        </ul>
                                        {runStopBeforeStartLogs.length > 5 && (
                                            <span>
                                                {renderAlertMessage('EPR_RUNSTOP_MORE', { Count: runStopBeforeStartLogs.length - 5 })}
                                            </span>
                                        )}
                                    </div>
                                )}

                                {runStopAfterEndLogs.length > 0 && (
                                    <div className={styles.runStopWarningSection}>
                                        <h4>
                                            {getAlertText('EPR_RUNSTOP_AFTER_END_TITLE')}
                                        </h4>
                                        <ul>
                                            {runStopAfterEndLogs.slice(0, 5).map((log: any, index: number) => {
                                                const {
                                                    timestampField
                                                } = getRunStopFieldMapping(selectedStation);

                                                const dateField = log[timestampField];
                                                return (
                                                    <li key={log.Id || index}>
                                                        {dateField}
                                                    </li>
                                                );
                                            })}
                                        </ul>
                                        {runStopAfterEndLogs.length > 5 && (
                                            <span>
                                                {renderAlertMessage('EPR_RUNSTOP_MORE', {
                                                    Count: runStopAfterEndLogs.length - 5
                                                })}
                                            </span>
                                        )}
                                    </div>
                                )}
                                <p>
                                    {getAlertText('EPR_SUBMIT_RUNSTOP_CONTINUE')}
                                </p>
                            </>
                        )}

                        {eprPopupType === 'timeError' && (
                            <>
                                <p>
                                    {renderAlertMessage('EPR_TIME_ERROR')}
                                </p>
                                {startTimeError && (
                                    <div className={styles.timeErrorMessage}>
                                        <strong>Start Time:</strong> {startTimeError}
                                    </div>
                                )}
                                {endTimeError && (
                                    <div className={styles.timeErrorMessage}>
                                        <strong>End Time:</strong> {endTimeError}
                                    </div>
                                )}
                            </>
                        )}

                        {eprPopupType === 'dataTimeError' && (
                            <p>
                                {renderAlertMessage('EPR_PERFORMANCE_DATA_ERROR')}
                            </p>
                        )}

                        {eprPopupType === 'closeShift' && (
                            <p>
                                {renderAlertMessage('EPR_CLOSE_SHIFT_CONFIRMATION')}
                            </p>
                        )}

                        {eprPopupType === 'updateRunStop' && (
                            <>
                                <p>
                                    {renderAlertMessage('EPR_UPDATE_RUNSTOP_WARNING')}
                                </p>

                                {runStopBeforeStartLogs.length > 0 && (
                                    <div className={styles.runStopWarningSection}>
                                        <h4>
                                            {getAlertText('EPR_RUNSTOP_BEFORE_START_TITLE')}
                                        </h4>
                                        <ul>
                                            {runStopBeforeStartLogs.slice(0, 5).map((log: any, index: number) => {
                                                const {
                                                    timestampField
                                                } = getRunStopFieldMapping(selectedStation);

                                                const dateField = log[timestampField];
                                                return (
                                                    <li key={log.Id || index}>
                                                        {dateField}
                                                    </li>
                                                );
                                            })}
                                        </ul>
                                        {runStopBeforeStartLogs.length > 5 && (
                                            <span>
                                                {renderAlertMessage('EPR_RUNSTOP_MORE', { Count: runStopBeforeStartLogs.length - 5 })}
                                            </span>
                                        )}
                                    </div>
                                )}

                                {runStopAfterEndLogs.length > 0 && (
                                    <div className={styles.runStopWarningSection}>
                                        <h4>
                                            {getAlertText('EPR_RUNSTOP_AFTER_END_TITLE')}
                                        </h4>
                                        <ul>
                                            {runStopAfterEndLogs.slice(0, 5).map((log: any, index: number) => {
                                                const {
                                                    timestampField
                                                } = getRunStopFieldMapping(selectedStation);

                                                const dateField = log[timestampField];
                                                return (
                                                    <li key={log.Id || index}>
                                                        {dateField}
                                                    </li>
                                                );
                                            })}
                                        </ul>
                                        {runStopAfterEndLogs.length > 5 && (
                                            <span>
                                                {renderAlertMessage('EPR_RUNSTOP_MORE', {
                                                    Count: runStopAfterEndLogs.length - 5
                                                })}
                                            </span>
                                        )}
                                    </div>
                                )}
                                <p>
                                    {renderAlertMessage('EPR_UPDATE_TIME_CONTINUE')}
                                </p>
                            </>
                        )}

                        {eprPopupType === 'updateTimeWarning' && (
                            <>
                                {timeUpdateShiftWarning === 'gap' && (
                                    <>
                                        <h3>
                                            {getAlertText('EPR_SHIFT_TIME_GAP_TITLE')}
                                        </h3>

                                        <p>
                                            {renderAlertMessage('EPR_GAP_WARNING', { PreviousShiftEndTime: previousShiftEndTime, CurrentStartTime: currentStartTime, Minutes: eprPopupMinutes })}
                                        </p>
                                    </>
                                )}
                                {timeUpdateShiftWarning === 'overlap' && (
                                    <>
                                        <h3>
                                            {getAlertText('EPR_SHIFT_TIME_OVERLAP_TITLE')}
                                        </h3>

                                        <p>
                                            {renderAlertMessage('EPR_OVERLAP_WARNING', { PreviousShiftEndTime: previousShiftEndTime, CurrentStartTime: currentStartTime, Minutes: eprPopupMinutes })}
                                        </p>
                                    </>
                                )}

                                {(runStopBeforeStartLogs.length > 0 || runStopAfterEndLogs.length > 0) && (
                                    <>
                                        <hr />
                                        <h3>
                                            {getAlertText('EPR_RUNSTOP_WARNING_TITLE')}
                                        </h3>
                                        <p>
                                            {getAlertText('EPR_RUNSTOP_WARNING')}
                                        </p>
                                        {runStopBeforeStartLogs.length > 0 && (
                                            <div className={styles.runStopWarningSection}>
                                                <h4>
                                                    {getAlertText('EPR_RUNSTOP_BEFORE_START_TITLE')}
                                                </h4>
                                                <ul>
                                                    {runStopBeforeStartLogs.slice(0, 5).map((log: any, index: number) => {
                                                        const {
                                                            timestampField
                                                        } = getRunStopFieldMapping(selectedStation);

                                                        const dateField = log[timestampField];
                                                        return (
                                                            <li key={log.Id || index}>
                                                                {dateField}
                                                            </li>
                                                        );
                                                    })}
                                                </ul>
                                                {runStopBeforeStartLogs.length > 5 && (
                                                    <span>
                                                        {renderAlertMessage('EPR_RUNSTOP_MORE', { Count: runStopBeforeStartLogs.length - 5 })}
                                                    </span>
                                                )}
                                            </div>
                                        )}

                                        {runStopAfterEndLogs.length > 0 && (
                                            <div className={styles.runStopWarningSection}>
                                                <h4>
                                                    {getAlertText('EPR_RUNSTOP_AFTER_END_TITLE')}
                                                </h4>
                                                <ul>
                                                    {runStopAfterEndLogs.slice(0, 5).map((log: any, index: number) => {
                                                        const { timestampField } = getRunStopFieldMapping(selectedStation);
                                                        const dateField = log[timestampField];

                                                        return (
                                                            <li key={log.Id || index}>
                                                                {dateField}
                                                            </li>
                                                        );
                                                    })}
                                                </ul>
                                                {runStopAfterEndLogs.length > 5 && (
                                                    <span>
                                                        {renderAlertMessage('EPR_RUNSTOP_MORE', { Count: runStopAfterEndLogs.length - 5 })}
                                                    </span>
                                                )}
                                            </div>
                                        )}
                                    </>
                                )}
                                <p>
                                    {getAlertText('EPR_UPDATE_TIME_CONTINUE')}
                                </p>
                            </>
                        )}

                    </div>
                </div>
                <div className={styles.popup_button_box}>
                    <button type="button" className={styles.popup_button_1} onClick={() => handleEPRPopupOK()}>
                        OK
                    </button>
                    <button type="button" className={styles.popup_button_2} onClick={() => handleEPRPopupCancel()}>
                        Cancel
                    </button>
                </div>
            </Modal>

            <Modal isOpen={showSuccessPopup} isBlocking={false}>
                <div className={styles.modalPopup_header}>
                    <h2>Success</h2>
                </div>

                <div className={styles.modalPopup}>
                    <div className={styles.note_box}>
                        <p>
                            {renderAlertMessage('EPR_SAVE_SUCCESS')}
                        </p>
                    </div>
                </div>

                <div className={styles.popup_button_box}>
                    <button type="button" className={styles.popup_button_1} onClick={() => setShowSuccessPopup(false)}>
                        OK
                    </button>
                </div>
            </Modal>

            <Modal isOpen={showApprovalPopup} isBlocking={true}>
                <div className={styles.modalPopup_header}>
                    <h2>Success</h2>
                </div>

                <div className={styles.modalPopup}>
                    <div className={styles.note_box}>
                        <p>
                            {renderAlertMessage('EPR_APPROVAL_SUCCESS', { ApprovalAction: approvalAction === 'Approved' ? 'approved' : 'rejected' })}
                        </p>
                    </div>
                </div>

                <div className={styles.popup_button_box}>
                    <button type="button" className={styles.popup_button_1} onClick={() => setShowApprovalPopup(false)}>
                        OK
                    </button>
                </div>
            </Modal>

        </div>
    );
};

export default EPRHeader;