import * as React from 'react';
import type { IErwMillProps } from './IErwMillProps';
import styles from '../css/Common.module.scss';
import EPRHeader from '../components/EPRHeader';
import PerformanceTable from '../components/PerformanceTable';
import { Modal } from '@fluentui/react';

const ErwMill: React.FC<IErwMillProps> = (props) => {

  const [loading, setLoading] = React.useState<boolean>(true);
  const [stations, setStations] = React.useState<any[]>([]);
  const [shiftHours, setShiftHours] = React.useState<any[]>([]);
  const [selectedStations, setSelectedStations] = React.useState<string>('ERW Mill');
  const [size, setSize] = React.useState<any[]>([]);
  const [grade, setGrade] = React.useState<any[]>([]);
  const [importantNotes, setImportantNotes] = React.useState<any[]>([]);

  const [byHours, setByHours] = React.useState<any[]>([]);
  const [classTypes, setClassTypes] = React.useState<any[]>([]);
  const [classTypes2, setClassTypes2] = React.useState<any[]>([]);
  const [classTypes3, setClassTypes3] = React.useState<any[]>([]);
  const [classTypes4, setClassTypes4] = React.useState<any[]>([]);
  const [leadsMember, setLeadsMember] = React.useState<any[]>([]);

  const [EPRAlertsMessages, setEPRAlertsMessages] = React.useState<any[]>([]);

  const [runStopLogs, setRunStopLogs] = React.useState<any[]>([]);
  const [currentStartTime, setCurrentStartTime] = React.useState<string>('');
  const [currentEndTime, setCurrentEndTime] = React.useState<string>('');
  const [currentEPRId, setCurrentEPRId] = React.useState<number | null>(null);
  const [currentEPRDate, setCurrentEPRDate] = React.useState<string>('');

  const [performanceRows, setPerformanceRows] = React.useState<any[]>([]);
  const [approvalErrorRowIds, setApprovalErrorRowIds] = React.useState<number[]>([]);
  const [isEPRReadOnly, setIsEPRReadOnly] = React.useState<boolean>(false);
  const [isPageLoading, setIsPageLoading] = React.useState<boolean>(false);

  const [showApprovalPopup, setShowApprovalPopup] = React.useState<boolean>(false);
  const [approvalPopupType, setApprovalPopupType] = React.useState<'error' | 'success' | 'submit'>('submit');

  const [showRefreshTimer, setShowRefreshTimer] = React.useState<boolean>(false);
  const [refreshSeconds, setRefreshSeconds] = React.useState<number>(120);
  const idleTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  const refreshTimerRef = React.useRef<ReturnType<typeof setInterval> | null>(null);
  const refreshTimeoutRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  const isRefreshingRef = React.useRef<boolean>(false);

  React.useEffect(() => {
    loadData();
  }, []);

  React.useEffect(() => {
    if (!currentEPRId) {
      setShowRefreshTimer(false);
      setRefreshSeconds(120);

      if (idleTimerRef.current) {
        clearTimeout(idleTimerRef.current);
      }

      if (refreshTimerRef.current) {
        clearInterval(refreshTimerRef.current);
      }

      if (refreshTimeoutRef.current) {
        clearTimeout(refreshTimeoutRef.current);
      }

      return;
    }

    const IDLE_TIME = 2 * 60 * 1000;
    const REFRESH_TIME = 2 * 60 * 1000;

    const clearAllTimers = (): void => {
      if (idleTimerRef.current) {
        clearTimeout(idleTimerRef.current);
        idleTimerRef.current = null;
      }

      if (refreshTimerRef.current) {
        clearInterval(refreshTimerRef.current);
        refreshTimerRef.current = null;
      }

      if (refreshTimeoutRef.current) {
        clearTimeout(refreshTimeoutRef.current);
        refreshTimeoutRef.current = null;
      }
    };

    const startIdleTimer = (): void => {
      if (idleTimerRef.current) {
        clearTimeout(idleTimerRef.current);
      }
      setShowRefreshTimer(false);
      setRefreshSeconds(120);
      idleTimerRef.current = setTimeout(() => {
        setShowRefreshTimer(true);
        setRefreshSeconds(120);
        refreshTimerRef.current = setInterval(() => {
          setRefreshSeconds((previous) => {
            if (previous <= 1) {
              if (refreshTimerRef.current) {
                clearInterval(refreshTimerRef.current);
                refreshTimerRef.current = null;
              }
              return 0;
            }
            return previous - 1;
          });
        }, 1000);

        refreshTimeoutRef.current = setTimeout(async () => {
          if (refreshTimerRef.current) {
            clearInterval(refreshTimerRef.current);
            refreshTimerRef.current = null;
          }
          setShowRefreshTimer(false);
          setRefreshSeconds(120);
          await refreshCurrentEPRData();
          startIdleTimer();
        }, REFRESH_TIME);
      }, IDLE_TIME);
    };

    const handleUserActivity = (): void => {
      clearAllTimers();
      startIdleTimer();
    };

    const activityEvents = ['pointermove', 'pointerdown', 'wheel', 'keydown', 'input', 'touchstart'];
    activityEvents.forEach((eventName) => {
      window.addEventListener(eventName, handleUserActivity, { passive: true });
    });

    startIdleTimer();

    return () => {
      clearAllTimers();
      activityEvents.forEach((eventName) => { window.removeEventListener(eventName, handleUserActivity); });
    };
  }, [currentEPRId, currentEPRDate, currentStartTime, currentEndTime, selectedStations]);

  const loadData = async (): Promise<void> => {
    try {
      setLoading(true);
      const [stationsData, sizeData, gradeData, importantNotes, EPRAlertsMessages] = await Promise.all([props.spService.getStations(), props.spService.getSize(), props.spService.getGrade(), props.spService.getImportantNotes(), props.spService.getEPRAlertsMessages()]);
      setStations(stationsData);
      setSelectedStations('ERW Mill');
      setSize(sizeData);
      setGrade(gradeData);
      setImportantNotes(importantNotes);
      setEPRAlertsMessages(EPRAlertsMessages);

      const selectedStationItem = stationsData.find((item: any) => item.Title === 'ERW Mill');
      if (selectedStationItem?.Id) {

        const hoursData = await props.spService.getShiftHours(selectedStationItem.Id);
        setShiftHours(hoursData);

        const classTypeData = await props.spService.getClass1(selectedStationItem.Id);
        setClassTypes(classTypeData);

        const classData = await props.spService.getClass2();
        setClassTypes2(classData);

        const VoiceTypeData = await props.spService.getClass3();
        setClassTypes3(VoiceTypeData);

        const VoiceData = await props.spService.getClass4();
        setClassTypes4(VoiceData);

        const LeadData = await props.spService.getParticipant(selectedStationItem.Id);
        setLeadsMember(LeadData);
      }
    }
    catch (error) {
      console.error('Error loading ERW Mill data:', error);
    }
    finally {
      setLoading(false);
    }
  };

  const handleEPRLoaded = async (
    eprId: number,
    station: string,
    date: string,
    shift: string,
    startTime: string,
    endTime: string
  ): Promise<void> => {
    try {
      setIsPageLoading(true);
      await getAllData(eprId, station, date, shift);
      setCurrentEPRId(eprId);
      setCurrentEPRDate(date);
      setCurrentStartTime(startTime);
      setCurrentEndTime(endTime);
    }
    catch (error) {
      console.error('Error loading EPR data:', error);
    }
    finally {
      setIsPageLoading(false);
    }
  };

  const getAllData = async (Id: number, station: string, date: string, shift: string): Promise<void> => {
    try {
      const byHoursData = await props.spService.getEPRByHours(station, Id);
      const selectedDateLogs = await props.spService.getEPRrunStopLogs(station, date);
      let allRunStopLogs = selectedDateLogs || [];

      if (shift === 'Shift 2') {
        const nextDate = new Date(`${date}T00:00:00`);
        nextDate.setDate(nextDate.getDate() + 1);
        const year = nextDate.getFullYear();
        const month = String(nextDate.getMonth() + 1).padStart(2, '0');
        const day = String(nextDate.getDate()).padStart(2, '0');
        const nextDateString = `${year}-${month}-${day}`;
        const nextDateLogs = await props.spService.getEPRrunStopLogs(station, nextDateString);
        allRunStopLogs = [...allRunStopLogs, ...(nextDateLogs || [])];
      }

      const uniqueRunStopLogs = Array.from(new Map(allRunStopLogs.map((log: any) => [log.Id, log])).values());
      setByHours(byHoursData || []);
      setRunStopLogs(uniqueRunStopLogs);
    }
    catch (error) {
      console.error('Error refreshing EPR data:', error);
    }
  };

  const refreshCurrentEPRData = async (): Promise<void> => {
    if (!currentEPRId || !currentEPRDate || isRefreshingRef.current) {
      return;
    }

    try {
      isRefreshingRef.current = true;
      const shift = currentStartTime === '4:00 PM' && currentEndTime === '2:00 AM' ? 'Shift 2' : 'Shift 1';
      await getAllData(currentEPRId, selectedStations, currentEPRDate, shift);
    }
    catch (error) {
      console.error('Error during automatic refresh:', error);
    }
    finally {
      isRefreshingRef.current = false;
    }
  };

  const refreshByHours = async (): Promise<void> => {
    if (!currentEPRId) {
      return;
    }
    try {
      const data = await props.spService.getEPRByHours(selectedStations, currentEPRId);
      setByHours(data || []);
    }
    catch (error) {
      console.error('Error refreshing By Hours:', error);
    }
  };

  const handleSubmitForApproval = async (eprId: number, station: string): Promise<boolean> => {
    const incompleteRows = performanceRows.filter((row: any) => row.MachineState === 'Stopped' && (!row.Class1 || !row.Class2 || !row.Class3));
    if (incompleteRows.length > 0) {
      setApprovalErrorRowIds(incompleteRows.map((row: any) => row.id));
      setApprovalPopupType('submit');
      setShowApprovalPopup(true);
      return false;
    }
    setApprovalErrorRowIds([]);

    try {
      await props.spService.updateEPR(station, eprId, { IsClosed: true, SubmitForApproval: true, LevelOneApprovalStatus: "Pending" });
      setIsEPRReadOnly(true);
      setApprovalPopupType('success');
      setShowApprovalPopup(true);
      return true;
    }
    catch (error) {
      console.error('Error submitting EPR for approval:', error);
      setApprovalPopupType('error');
      setShowApprovalPopup(true);
      return false;
    }
  };

  return (
    <div className={styles.custom_page}>

      {!loading && (<>

        <EPRHeader
          stations={stations}
          SelectedStations={selectedStations}
          size={size}
          grade={grade}
          leadsMember={leadsMember}
          spService={props.spService}
          sp={props.spService.sp}
          hasPerformanceData={byHours.length > 0}
          performanceRows={performanceRows}
          shiftHours={shiftHours}
          importantNotes={importantNotes}
          EPRAlertsMessages={EPRAlertsMessages}
          onEPRLoaded={handleEPRLoaded}
          onSubmitForApproval={handleSubmitForApproval}
          onReadOnlyChange={setIsEPRReadOnly}
          onLoadingChange={setIsPageLoading}
          onTimeChange={(startTime, endTime) => { setCurrentStartTime(startTime); setCurrentEndTime(endTime); }}
        />

        {currentEPRId && (
          <PerformanceTable byHours={byHours} runStopLogs={runStopLogs} startTime={currentStartTime} endTime={currentEndTime} classTypes={classTypes}
            classTypes2={classTypes2}
            classTypes3={classTypes3}
            classTypes4={classTypes4}
            spService={props.spService}
            SelectedStations={selectedStations} eprId={currentEPRId} eprDate={currentEPRDate} approvalErrorRowIds={approvalErrorRowIds}
            onByHoursRefresh={refreshByHours} onRowsChange={setPerformanceRows} isEPRReadOnly={isEPRReadOnly}
          />
        )}

        {currentEPRId && showRefreshTimer && (
          <div className={styles.autoRefreshTimer}>
            <svg className={styles.autoRefreshCircle} viewBox="0 0 100 100">
              <circle className={styles.autoRefreshCircleBg} cx="50" cy="50" r="44" />
              <circle className={styles.autoRefreshCircleProgress} cx="50" cy="50" r="44" style={{ strokeDashoffset: 276.46 - (276.46 * refreshSeconds) / 120 }} />
            </svg>
            <div className={styles.autoRefreshTimerText}>
              {String(Math.floor(refreshSeconds / 60)).padStart(2, '0')}:{String(refreshSeconds % 60).padStart(2, '0')}
            </div>
          </div>
        )}

      </>)}

      {(loading || isPageLoading) && (
        <div className={styles.chart_loading_box}>
          <div className={styles.chart_loading}></div>
        </div>
      )}

      <Modal isOpen={showApprovalPopup} isBlocking={false}>
        <div className={styles.modalPopup_header}>
          <h2>Equipment Performance Report</h2>
        </div>

        <div className={styles.modalPopup}>
          <div className={styles.note_box}>
            {approvalPopupType === 'submit' && (
              <>
                <p>
                  Please complete all required fields for the stopped entries before submitting for approval.
                </p>
                <p>
                  The incomplete entries have been highlighted in red.
                </p>
              </>
            )}

            {approvalPopupType === 'success' && (
              <p>EPR has been submitted for approval successfully.</p>
            )}

            {approvalPopupType === 'error' && (
              <p>Unable to submit the EPR for approval.</p>
            )}
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

export default ErwMill;