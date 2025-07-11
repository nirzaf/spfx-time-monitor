import { sp } from '@pnp/sp/presets/all';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ILeaveRequest, ILeaveType, ILeaveBalance } from '../models/ILeaveModels';

/**
 * SharePoint service for managing leave requests, types, and balances
 */
export class SharePointService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
    sp.setup({
      spfxContext: context
    });
  }

  /**
   * Get all leave types
   */
  public async getLeaveTypes(): Promise<ILeaveType[]> {
    try {
      const items = await sp.web.lists.getByTitle('Leave Types').items
        .select('ID', 'Title', 'Description', 'IsActive', 'RequiresApproval', 'MaxDaysPerRequest', 'RequiresDocumentation', 'ColorCode', 'PolicyURL')
        .filter('IsActive eq true')
        .get();
      
      return items.map(item => ({
        id: item.ID,
        title: item.Title,
        description: item.Description,
        isActive: item.IsActive,
        requiresApproval: item.RequiresApproval,
        maxDaysPerRequest: item.MaxDaysPerRequest,
        requiresDocumentation: item.RequiresDocumentation,
        colorCode: item.ColorCode,
        policyURL: item.PolicyURL
      }));
    } catch (error) {
      console.error('Error fetching leave types:', error);
      throw error;
    }
  }

  /**
   * Create a new leave request
   */
  public async createLeaveRequest(request: Partial<ILeaveRequest>): Promise<ILeaveRequest> {
    try {
      const result = await sp.web.lists.getByTitle('Leave Requests').items.add({
        Title: `${request.leaveType} - ${request.startDate}`,
        RequesterId: request.requesterId,
        EmployeeID: request.employeeId,
        Department: request.department,
        ManagerId: request.managerId,
        LeaveTypeId: request.leaveTypeId,
        StartDate: request.startDate,
        EndDate: request.endDate,
        TotalDays: request.totalDays,
        IsPartialDay: request.isPartialDay,
        PartialDayHours: request.partialDayHours,
        RequestComments: request.requestComments,
        ApprovalStatus: 'Pending',
        SubmissionDate: new Date().toISOString(),
        NotificationsSent: false
      });

      return {
        id: result.data.ID,
        title: result.data.Title,
        requesterId: result.data.RequesterId,
        employeeId: result.data.EmployeeID,
        department: result.data.Department,
        managerId: result.data.ManagerId,
        leaveTypeId: result.data.LeaveTypeId,
        leaveType: request.leaveType,
        startDate: result.data.StartDate,
        endDate: result.data.EndDate,
        totalDays: result.data.TotalDays,
        isPartialDay: result.data.IsPartialDay,
        partialDayHours: result.data.PartialDayHours,
        requestComments: result.data.RequestComments,
        approvalStatus: result.data.ApprovalStatus,
        submissionDate: result.data.SubmissionDate,
        approvalDate: result.data.ApprovalDate,
        approvalComments: result.data.ApprovalComments
      };
    } catch (error) {
      console.error('Error creating leave request:', error);
      throw error;
    }
  }

  /**
   * Get leave requests for current user
   */
  public async getUserLeaveRequests(userId: number): Promise<ILeaveRequest[]> {
    try {
      const items = await sp.web.lists.getByTitle('Leave Requests').items
        .select('*', 'LeaveType/Title', 'Requester/Title', 'Manager/Title')
        .expand('LeaveType', 'Requester', 'Manager')
        .filter(`RequesterId eq ${userId}`)
        .orderBy('SubmissionDate', false)
        .get();

      return items.map(item => ({
        id: item.ID,
        title: item.Title,
        requesterId: item.RequesterId,
        employeeId: item.EmployeeID,
        department: item.Department,
        managerId: item.ManagerId,
        leaveTypeId: item.LeaveTypeId,
        leaveType: item.LeaveType?.Title,
        startDate: item.StartDate,
        endDate: item.EndDate,
        totalDays: item.TotalDays,
        isPartialDay: item.IsPartialDay,
        partialDayHours: item.PartialDayHours,
        requestComments: item.RequestComments,
        approvalStatus: item.ApprovalStatus,
        submissionDate: item.SubmissionDate,
        approvalDate: item.ApprovalDate,
        approvalComments: item.ApprovalComments
      }));
    } catch (error) {
      console.error('Error fetching user leave requests:', error);
      throw error;
    }
  }

  /**
   * Get all leave requests for calendar view
   */
  public async getAllLeaveRequests(): Promise<ILeaveRequest[]> {
    try {
      const items = await sp.web.lists.getByTitle('Leave Requests').items
        .select('*', 'LeaveType/Title', 'LeaveType/ColorCode', 'Requester/Title', 'Manager/Title')
        .expand('LeaveType', 'Requester', 'Manager')
        .filter('ApprovalStatus eq \'Approved\'')
        .get();

      return items.map(item => ({
        id: item.ID,
        title: item.Title,
        requesterId: item.RequesterId,
        employeeId: item.EmployeeID,
        department: item.Department,
        managerId: item.ManagerId,
        leaveTypeId: item.LeaveTypeId,
        leaveType: item.LeaveType?.Title,
        startDate: item.StartDate,
        endDate: item.EndDate,
        totalDays: item.TotalDays,
        isPartialDay: item.IsPartialDay,
        partialDayHours: item.PartialDayHours,
        requestComments: item.RequestComments,
        approvalStatus: item.ApprovalStatus,
        submissionDate: item.SubmissionDate,
        approvalDate: item.ApprovalDate,
        approvalComments: item.ApprovalComments,
        requesterName: item.Requester?.Title,
        colorCode: item.LeaveType?.ColorCode
      }));
    } catch (error) {
      console.error('Error fetching all leave requests:', error);
      throw error;
    }
  }

  /**
   * Get leave balance for user
   */
  public async getUserLeaveBalance(userId: number, leaveTypeId: number): Promise<ILeaveBalance | null> {
    try {
      const items = await sp.web.lists.getByTitle('Leave Balances').items
        .select('*', 'LeaveType/Title')
        .expand('LeaveType')
        .filter(`EmployeeId eq ${userId} and LeaveTypeId eq ${leaveTypeId}`)
        .get();

      if (items.length > 0) {
        const item = items[0];
        return {
          id: item.ID,
          employeeId: item.EmployeeId,
          leaveTypeId: item.LeaveTypeId,
          leaveType: item.LeaveType?.Title,
          totalAllowance: item.TotalAllowance,
          usedDays: item.UsedDays,
          remainingDays: item.RemainingDays,
          carryOverDays: item.CarryOverDays,
          effectiveDate: item.EffectiveDate,
          expirationDate: item.ExpirationDate
        };
      }
      return null;
    } catch (error) {
      console.error('Error fetching leave balance:', error);
      throw error;
    }
  }

  /**
   * Update leave request status
   */
  public async updateLeaveRequestStatus(requestId: number, status: string, comments?: string): Promise<void> {
    try {
      const updateData: any = {
        ApprovalStatus: status,
        ApprovalDate: new Date().toISOString()
      };

      if (comments) {
        updateData.ApprovalComments = comments;
      }

      await sp.web.lists.getByTitle('Leave Requests').items.getById(requestId).update(updateData);
    } catch (error) {
      console.error('Error updating leave request status:', error);
      throw error;
    }
  }

  /**
   * Get current user information
   */
  public async getCurrentUser(): Promise<{ id: number; title: string; email: string }> {
    try {
      const user = await sp.web.currentUser.get();
      return {
        id: user.Id,
        title: user.Title,
        email: user.Email
      };
    } catch (error) {
      console.error('Error fetching current user:', error);
      throw error;
    }
  }
}