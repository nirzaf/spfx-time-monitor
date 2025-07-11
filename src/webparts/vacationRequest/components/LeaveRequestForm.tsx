import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  PrimaryButton,
  DefaultButton,
  TextField,
  Dropdown,
  IDropdownOption,
  DatePicker,
  Checkbox,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Panel,
  PanelType
} from '@fluentui/react';
import { SharePointService } from '../../../services/SharePointService';
import { ILeaveType, ILeaveRequest, ILeaveBalance } from '../../../models/ILeaveModels';
import styles from './LeaveRequestForm.module.scss';
import type { ILeaveRequestFormProps } from './ILeaveRequestFormProps';

/**
 * State interface for the Leave Request Form component
 */
interface ILeaveRequestFormState {
  leaveTypes: ILeaveType[];
  selectedLeaveType: number | null;
  startDate: Date | null;
  endDate: Date | null;
  isPartialDay: boolean;
  partialDayHours: number;
  comments: string;
  isLoading: boolean;
  isSubmitting: boolean;
  message: { text: string; type: MessageBarType } | null;
  leaveBalance: ILeaveBalance | null;
  totalDays: number;
  showSuccessPanel: boolean;
}

/**
 * LeaveRequestForm React functional component
 * Provides a comprehensive form for submitting leave requests
 */
const LeaveRequestForm: React.FunctionComponent<ILeaveRequestFormProps> = (props) => {
  const {
    context,
    title,
    description,
    hasTeamsContext,
    userDisplayName
  } = props;

  // Initialize SharePoint service
  const [spService] = useState(() => new SharePointService(context));

  // Component state
  const [state, setState] = useState<ILeaveRequestFormState>({
    leaveTypes: [],
    selectedLeaveType: null,
    startDate: null,
    endDate: null,
    isPartialDay: false,
    partialDayHours: 8,
    comments: '',
    isLoading: true,
    isSubmitting: false,
    message: null,
    leaveBalance: null,
    totalDays: 0,
    showSuccessPanel: false
  });

  /**
   * Load leave types on component mount
   */
  useEffect(() => {
    loadLeaveTypes();
  }, []);

  /**
   * Calculate total days when dates change
   */
  useEffect(() => {
    if (state.startDate && state.endDate) {
      calculateTotalDays();
    }
  }, [state.startDate, state.endDate, state.isPartialDay, state.partialDayHours]);

  /**
   * Load leave balance when leave type changes
   */
  useEffect(() => {
    if (state.selectedLeaveType) {
      loadLeaveBalance();
    }
  }, [state.selectedLeaveType]);

  /**
   * Load available leave types from SharePoint
   */
  const loadLeaveTypes = async (): Promise<void> => {
    try {
      const leaveTypes = await spService.getLeaveTypes();
      setState(prev => ({
        ...prev,
        leaveTypes,
        isLoading: false
      }));
    } catch (error) {
      setState(prev => ({
        ...prev,
        isLoading: false,
        message: {
          text: 'Failed to load leave types. Please refresh the page.',
          type: MessageBarType.error
        }
      }));
    }
  };

  /**
   * Load leave balance for selected leave type
   */
  const loadLeaveBalance = async (): Promise<void> => {
    if (!state.selectedLeaveType) return;

    try {
      const userId = context.pageContext.legacyPageContext.userId;
      const balance = await spService.getUserLeaveBalance(userId, state.selectedLeaveType);
      setState(prev => ({
        ...prev,
        leaveBalance: balance
      }));
    } catch (error) {
      console.error('Error loading leave balance:', error);
    }
  };

  /**
   * Calculate total days for the leave request
   */
  const calculateTotalDays = (): void => {
    if (!state.startDate || !state.endDate) return;

    const start = new Date(state.startDate.toISOString());
    const end = new Date(state.endDate.toISOString());
    const timeDiff = end.getTime() - start.getTime();
    let daysDiff = Math.ceil(timeDiff / (1000 * 3600 * 24)) + 1;

    if (state.isPartialDay && daysDiff === 1) {
      daysDiff = state.partialDayHours / 8;
    }

    setState(prev => ({
      ...prev,
      totalDays: daysDiff
    }));
  };

  /**
   * Handle form submission
   */
  const handleSubmit = async (): Promise<void> => {
    if (!validateForm()) return;

    setState(prev => ({ ...prev, isSubmitting: true, message: null }));

    try {
      const userId = context.pageContext.legacyPageContext.userId;
      const userProfile = context.pageContext.user;
      
      const leaveRequest: Partial<ILeaveRequest> = {
        requesterId: userId,
        employeeId: userProfile.loginName,
        department: 'IT', // This should come from user profile
        managerId: 1, // This should come from user profile
        leaveTypeId: state.selectedLeaveType!,
        leaveType: state.leaveTypes.filter(lt => lt.id === state.selectedLeaveType)[0]?.title,
        startDate: state.startDate!.toISOString(),
        endDate: state.endDate!.toISOString(),
        totalDays: state.totalDays,
        isPartialDay: state.isPartialDay,
        partialDayHours: state.isPartialDay ? state.partialDayHours : undefined,
        requestComments: state.comments
      };

      await spService.createLeaveRequest(leaveRequest);
      
      setState(prev => ({
        ...prev,
        isSubmitting: false,
        showSuccessPanel: true,
        message: {
          text: 'Leave request submitted successfully!',
          type: MessageBarType.success
        }
      }));

      // Reset form
      resetForm();
    } catch (error) {
      setState(prev => ({
        ...prev,
        isSubmitting: false,
        message: {
          text: 'Failed to submit leave request. Please try again.',
          type: MessageBarType.error
        }
      }));
    }
  };

  /**
   * Validate form data
   */
  const validateForm = (): boolean => {
    if (!state.selectedLeaveType) {
      setState(prev => ({
        ...prev,
        message: { text: 'Please select a leave type.', type: MessageBarType.error }
      }));
      return false;
    }

    if (!state.startDate || !state.endDate) {
      setState(prev => ({
        ...prev,
        message: { text: 'Please select start and end dates.', type: MessageBarType.error }
      }));
      return false;
    }

    if (state.startDate > state.endDate) {
      setState(prev => ({
        ...prev,
        message: { text: 'End date must be after start date.', type: MessageBarType.error }
      }));
      return false;
    }

    if (state.leaveBalance && state.totalDays > state.leaveBalance.remainingDays!) {
      setState(prev => ({
        ...prev,
        message: { text: 'Insufficient leave balance.', type: MessageBarType.error }
      }));
      return false;
    }

    return true;
  };

  /**
   * Reset form to initial state
   */
  const resetForm = (): void => {
    setState(prev => ({
      ...prev,
      selectedLeaveType: null,
      startDate: null,
      endDate: null,
      isPartialDay: false,
      partialDayHours: 8,
      comments: '',
      totalDays: 0,
      leaveBalance: null
    }));
  };

  /**
   * Handle leave type selection
   */
  const handleLeaveTypeChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option) {
      setState(prev => ({
        ...prev,
        selectedLeaveType: option.key as number,
        message: null
      }));
    }
  };

  /**
   * Prepare dropdown options for leave types
   */
  const getLeaveTypeOptions = (): IDropdownOption[] => {
    return state.leaveTypes.map(leaveType => ({
      key: leaveType.id!,
      text: leaveType.title!,
      data: leaveType
    }));
  };

  if (state.isLoading) {
    return (
      <div className={styles.loadingContainer}>
        <Spinner size={SpinnerSize.large} label="Loading leave request form..." />
      </div>
    );
  }

  return (
    <section className={`${styles.leaveRequestForm} ${hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.container}>
        <div className={styles.header}>
          <h2 className={styles.title}>{title || 'Leave Request Form'}</h2>
          <p className={styles.description}>{description}</p>
          <p className={styles.welcome}>Welcome, {userDisplayName}!</p>
        </div>

        {state.message && (
          <MessageBar
            messageBarType={state.message.type}
            onDismiss={() => setState(prev => ({ ...prev, message: null }))}
            dismissButtonAriaLabel="Close"
          >
            {state.message.text}
          </MessageBar>
        )}

        <div className={styles.formContainer}>
          <div className={styles.formSection}>
            <Dropdown
              label="Leave Type"
              placeholder="Select a leave type"
              options={getLeaveTypeOptions()}
              selectedKey={state.selectedLeaveType}
              onChange={handleLeaveTypeChange}
              required
              className={styles.formField}
            />

            {state.leaveBalance && (
              <div className={styles.balanceInfo}>
                <p><strong>Available Balance:</strong> {state.leaveBalance.remainingDays} days</p>
                <p><strong>Total Allowance:</strong> {state.leaveBalance.totalAllowance} days</p>
                <p><strong>Used:</strong> {state.leaveBalance.usedDays} days</p>
              </div>
            )}

            <div className={styles.dateSection}>
              <DatePicker
                label="Start Date"
                placeholder="Select start date"
                value={state.startDate || undefined}
                onSelectDate={(date) => setState(prev => ({ ...prev, startDate: date || null, message: null }))}
                isRequired
                className={styles.dateField}
              />

              <DatePicker
                label="End Date"
                placeholder="Select end date"
                value={state.endDate || undefined}
                onSelectDate={(date) => setState(prev => ({ ...prev, endDate: date || null, message: null }))}
                isRequired
                className={styles.dateField}
              />
            </div>

            <Checkbox
              label="Partial Day Leave"
              checked={state.isPartialDay}
              onChange={(ev, checked) => setState(prev => ({ ...prev, isPartialDay: !!checked }))}
              className={styles.formField}
            />

            {state.isPartialDay && (
              <TextField
                label="Hours"
                type="number"
                value={state.partialDayHours.toString()}
                onChange={(ev, value) => setState(prev => ({ ...prev, partialDayHours: parseFloat(value || '0') }))}
                min={0.5}
                max={8}
                step={0.5}
                className={styles.formField}
              />
            )}

            {state.totalDays > 0 && (
              <div className={styles.totalDays}>
                <strong>Total Days: {state.totalDays}</strong>
              </div>
            )}

            <TextField
              label="Comments (Optional)"
              multiline
              rows={4}
              value={state.comments}
              onChange={(ev, value) => setState(prev => ({ ...prev, comments: value || '' }))}
              placeholder="Add any additional comments or notes..."
              className={styles.formField}
            />

            <div className={styles.buttonSection}>
              <PrimaryButton
                text="Submit Request"
                onClick={handleSubmit}
                disabled={state.isSubmitting || !state.selectedLeaveType || !state.startDate || !state.endDate}
                className={styles.submitButton}
              />
              <DefaultButton
                text="Reset"
                onClick={resetForm}
                disabled={state.isSubmitting}
                className={styles.resetButton}
              />
            </div>
          </div>
        </div>

        <Panel
          isOpen={state.showSuccessPanel}
          onDismiss={() => setState(prev => ({ ...prev, showSuccessPanel: false }))}
          type={PanelType.medium}
          headerText="Request Submitted Successfully"
        >
          <div className={styles.successPanel}>
            <p>Your leave request has been submitted and is pending approval.</p>
            <p>You will receive an email notification once your manager reviews the request.</p>
            <PrimaryButton
              text="Close"
              onClick={() => setState(prev => ({ ...prev, showSuccessPanel: false }))}
            />
          </div>
        </Panel>
      </div>
    </section>
  );
};

export default LeaveRequestForm;
