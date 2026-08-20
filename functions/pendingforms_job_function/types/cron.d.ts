/**
 * Type of Cron Details object. Contains the details of the scheduled cron
 */
export interface CronDetails {
    getCronParam: (name?: string) => string;
    getAllCronParam?: () => Record<string, string>;
    getAllCronParams?: () => Record<string, string>;
    getRemainingExecutionCount: () => number;
    getCronDetails: () => Record<string, unknown>;
    getProjectDetails: () => Record<string, unknown>;
}

export interface Context {
    catalystHeaders: Record<string, string>;
    closeWithSuccess: () => void;
    closeWithFailure: () => void;
    getRemainingExecutionTimeMs: () => number;
    getMaxExecutionTimeMs: () => number;
}
