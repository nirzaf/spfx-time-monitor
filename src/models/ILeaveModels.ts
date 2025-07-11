/**
 * Interface for Leave Request data model
 */
export interface ILeaveRequest {
  id?: number;
  title?: string;
  requesterId?: number;
  employeeId?: string;
  department?: string;
  managerId?: number;
  leaveTypeId?: number;
  leaveType?: string;
  startDate?: string;
  endDate?: string;
  totalDays?: number;
  isPartialDay?: boolean;
  partialDayHours?: number;
  requestComments?: string;
  approvalStatus?: 'Pending' | 'Approved' | 'Rejected' | 'Cancelled';
  approvalDate?: string;
  approvalComments?: string;
  submissionDate?: string;
  lastModified?: string;
  attachmentURL?: string;
  workflowInstanceID?: string;
  notificationsSent?: boolean;
  calendarEventID?: string;
  requesterName?: string;
  managerName?: string;
  colorCode?: string;
}

/**
 * Interface for Leave Type data model
 */
export interface ILeaveType {
  id?: number;
  title?: string;
  description?: string;
  isActive?: boolean;
  requiresApproval?: boolean;
  maxDaysPerRequest?: number;
  requiresDocumentation?: boolean;
  colorCode?: string;
  policyURL?: string;
  createdDate?: string;
  modifiedDate?: string;
}

/**
 * Interface for Leave Balance data model
 */
export interface ILeaveBalance {
  id?: number;
  employeeId?: number;
  leaveTypeId?: number;
  leaveType?: string;
  totalAllowance?: number;
  usedDays?: number;
  remainingDays?: number;
  carryOverDays?: number;
  effectiveDate?: string;
  expirationDate?: string;
}

/**
 * Interface for Calendar Event data model
 */
export interface ICalendarEvent {
  id?: string;
  title?: string;
  start?: string;
  end?: string;
  backgroundColor?: string;
  borderColor?: string;
  textColor?: string;
  extendedProps?: {
    leaveRequestId?: number;
    employeeName?: string;
    leaveType?: string;
    status?: string;
    department?: string;
  };
}

/**
 * Interface for User Profile data
 */
export interface IUserProfile {
  id?: number;
  displayName?: string;
  email?: string;
  department?: string;
  managerId?: number;
  managerDisplayName?: string;
  employeeId?: string;
}

/**
 * Interface for Leave Request Form data
 */
export interface ILeaveRequestForm {
  leaveTypeId: number;
  startDate: Date;
  endDate: Date;
  isPartialDay: boolean;
  partialDayHours?: number;
  comments?: string;
  attachmentURL?: string;
}

/**
 * Interface for Leave Statistics
 */
export interface ILeaveStatistics {
  totalRequests?: number;
  pendingRequests?: number;
  approvedRequests?: number;
  rejectedRequests?: number;
  totalDaysRequested?: number;
  totalDaysApproved?: number;
  averageProcessingTime?: number;
}

/**
 * Interface for Team Coverage data
 */
export interface ITeamCoverage {
  date?: string;
  totalEmployees?: number;
  employeesOnLeave?: number;
  coveragePercentage?: number;
  criticalCoverage?: boolean;
}