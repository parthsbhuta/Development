import { SPFI } from '@pnp/sp';

export interface ILeftNavigation { Id: string; Title: string; IconColor: string; IconBW: string; URL: any; }
export interface IEPRListNames { epr: string; byHours: string; runStopLogs: string; }

export default class SharePointService {

    constructor(public sp: SPFI) { }

    public async getStations(): Promise<any[]> {
        const items = await this.sp.web.lists.getByTitle("Stations").items.select("Id", "Title", "URL").filter("IsActive eq 1").orderBy("Title", true)();
        return items;
    }

    public async getShiftHours(stationId: number): Promise<any[]> {
        const items = await this.sp.web.lists.getByTitle("Shift Hours").items.select("StartTime", "EndTime", "ShiftType", "BeforeStartWindowTime", "AfterEndWindowTime", 'Stations/Id', 'Stations/Title').expand('Stations').filter(`Stations/Id eq ${stationId}`)();
        return items;
    }

    public async getSize(): Promise<any[]> {
        const items = await this.sp.web.lists.getByTitle("Size").items.select("Id", "Title").filter("IsActive eq 1").orderBy("Title", true)();
        return items;
    }

    public async getGrade(): Promise<any[]> {
        const items = await this.sp.web.lists.getByTitle("Grade").items.select("Id", "Title").filter("IsActive eq 1").orderBy("Title", true)();
        return items;
    }

    public async getImportantNotes(): Promise<any[]> {
        const items = await this.sp.web.lists.getByTitle("Important Notes").items.top(1)();
        return items;
    }

    public async getEPRAlertsMessages(): Promise<any[]> {
        const items = await this.sp.web.lists.getByTitle("EPR Alerts Messages").items.filter("ForNewIntranet eq 1")();
        return items;
    }

    public async getParticipant(stationId: number): Promise<any[]> {
        const items = await this.sp.web.lists.getByTitle("Participant List").items.select("Id", "ApproverName/Title", "ApproverName/Id", "ApproverName/EMail", 'Stations/Id', 'Stations/Title').expand('ApproverName', 'Stations').filter(`Stations/Id eq ${stationId} and IsLead eq 1`)();
        return items;
    }

    public async getClass1(stationId: number): Promise<any[]> {
        const classTypes = await this.sp.web.lists.getByTitle('Class Type').items.select('*', 'Station/Id', 'Station/Title').expand('Station').filter(`Station/Id eq ${stationId}`)();
        return classTypes;
    }

    public async getClass2(): Promise<any[]> {
        const items = await this.sp.web.lists.getByTitle('Class').items.select('Id', '*', 'Title', 'ClassType/Id').expand('ClassType').top(1000)();
        return items;
    }

    public async getClass3(): Promise<any[]> {
        const items = await this.sp.web.lists.getByTitle('Voice Type').items.select('Id', '*', 'Title', 'Class/Id').expand('Class').top(1000)();
        return items;
    }

    public async getClass4(): Promise<any[]> {
        const items = await this.sp.web.lists.getByTitle('Voice').items.select('Id', '*', 'Title', 'VoiceType/Id').expand('VoiceType').top(1000)();
        return items;
    }

    private getEPRListNames(station: string): IEPRListNames | null {
        switch (station) {
            case 'Heat Treat': return { epr: 'EPRHeatTreatEquipmentPerformanceReportList', byHours: 'EPRHeatTreatByHours', runStopLogs: 'HTRunStopTimeLogs' };
            case 'Hydro': return { epr: 'EPRHydroEquipmentPerformanceReportList', byHours: 'EPRHydroByHours', runStopLogs: 'HydroRunStopTimeLogs' };
            case 'NDT': return { epr: 'EPRNDTEquipmentPerformanceReportList', byHours: 'EPRNDTByHours', runStopLogs: 'NDTRunStopTimeLogs' };
            case 'ERW Mill': return { epr: 'NewEPRMillEquipmentPerformanceReportList', byHours: 'NewEPRMillByHours', runStopLogs: 'MillRunStopTimeLogs' };
            case 'Threadline 1': return { epr: 'EPRThreadline1EquipmentPerformanceReportList', byHours: 'EPRThreadline1ByHours', runStopLogs: 'Threadline1RunStopTimeLogs' };
            case 'Threadline 2': return { epr: 'EPRThreadline2EquipmentPerformanceReportList', byHours: 'EPRThreadline2ByHours', runStopLogs: 'Threadline2RunStopTimeLogs' };
            default: return null;
        }
    }

    public async getEPRByDate(station: string, date: string, shifts: string): Promise<any[]> {
        const listName = this.getEPRListNames(station);
        if (!listName) {
            return [];
        }
        const isoDate = new Date(date).toISOString().split('T')[0];
        const start = `${isoDate}T00:00:00Z`;
        const end = `${isoDate}T23:59:00Z`;
        const filter = `Date ge '${start}' and ` + `Date le '${end}' and ` + `Shifts eq '${shifts}'`;
        const items = await this.sp.web.lists.getByTitle(listName.epr).items.select('*', 'LevelOneApprover/Title', 'LevelOneApprover/EMail').expand('LevelOneApprover').filter(filter)();
        return items;
    }

    public async getEPRByHours(station: string, eprId: number): Promise<any[]> {
        const listName = this.getEPRListNames(station);
        if (!listName) {
            throw new Error(`EPR list not found for station: ${station}`);
        }
        return await this.sp.web.lists.getByTitle(listName.byHours).items.filter(`EPR_ID eq ${eprId}`).select('*', 'ClassType/Id', 'ClassType/ViewTitle', 'Class/Id', 'Class/ViewTitle', 'VoiceType/Id', 'VoiceType/ViewTitle', 'Voice/Id', 'Voice/ViewTitle').expand('ClassType', 'Class', 'VoiceType', 'Voice')();
    }

    // public async getEPRrunStopLogs(station: string, date: string): Promise<any[]> {
    //     const listName = this.getEPRListNames(station);
    //     if (!listName) {
    //         throw new Error(`EPR list not found for station: ${station}`);
    //     }
    //     const [year, month, day] = date.split('-');
    //     const formattedDate = `${month}/${day}/${year}`;
    //     const filter = `field_10 eq '${formattedDate}'`;
    //     const items = await this.sp.web.lists.getByTitle(listName.runStopLogs).items.filter(filter).top(200)();
    //     return items;
    // }

    public async getEPRrunStopLogs(station: string, date: string): Promise<any[]> {
        const listName = this.getEPRListNames(station);
        if (!listName) {
            throw new Error(`EPR list not found for station: ${station}`);
        }

        const [year, month, day] = date.split('-');
        const formattedDate = `${month}/${day}/${year}`;

        let filter: string;

        if (station === 'Heat Treat' || station === 'ERW Mill') {
            filter = `field_10 eq '${formattedDate}'`;
        }
        else {
            filter = `field_3 eq '${formattedDate}'`;
        }

        const items = await this.sp.web.lists.getByTitle(listName.runStopLogs).items.filter(filter).top(200)();
        return items;
    }

    public async createEPR(station: string, data: any): Promise<any> {
        const listName = this.getEPRListNames(station);
        if (!listName) {
            throw new Error(`EPR list not found for station: ${station}`);
        }
        const result = await this.sp.web.lists.getByTitle(listName.epr).items.add(data);
        return result;
    }

    public async updateEPR(station: string, itemId: number, data: any): Promise<void> {
        const listName = this.getEPRListNames(station);
        if (!listName) {
            throw new Error(`EPR list not found for station: ${station}`);
        }
        await this.sp.web.lists.getByTitle(listName.epr).items.getById(itemId).update(data);
    }

    public async addEPRByHours(station: string, data: any): Promise<any> {
        const listName = this.getEPRListNames(station);
        if (!listName) {
            throw new Error(`EPR list not found for station: ${station}`);
        }
        const result = await this.sp.web.lists.getByTitle(listName.byHours).items.add(data);
        return result;
    }

    public async updateEPRByHours(station: string, id: number, data: any): Promise<any> {
        const listName = this.getEPRListNames(station);
        if (!listName) {
            throw new Error(`EPR list not found for station: ${station}`);
        }
        return await this.sp.web.lists.getByTitle(listName.byHours).items.getById(id).update(data);
    }

    public async deleteEPRByHours(station: string, itemId: number): Promise<void> {
        const listName = this.getEPRListNames(station);
        if (!listName) {
            throw new Error(`EPR list not found for station: ${station}`);
        }
        await this.sp.web.lists.getByTitle(listName.byHours).items.getById(itemId).delete();
    }

    public async getEPRById(station: string, itemId: number): Promise<any> {
        const listNames = this.getEPRListNames(station);
        if (!listNames) {
            return [];
        }
        return await this.sp.web.lists.getByTitle(listNames.epr).items.getById(itemId).select('*', 'Station/Id', 'Station/Title').expand('Station')();
    }

}
