import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import {
  Stack,
  Text,
  CommandBar,
  ICommandBarItemProps,
  DetailsList,
  IColumn,
  SelectionMode,
  Selection,
  Dropdown,
  IDropdownOption,
  TextField,
  DatePicker,
  Panel,
  PanelType,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Dialog,
  DialogType,
  DialogFooter,
  SearchBox,
  TooltipHost,
  Icon,
  Pivot,
  PivotItem,
  Label,
  ProgressIndicator
} from '@fluentui/react';
import { ILeaveAdministrationProps } from './ILeaveAdministrationProps';
import styles from '../LeaveAdministration.module.scss';
import { SharePointService } from '../../../services/SharePointService';
import {
  ILeaveRequest,
  ILeaveType,
  IUserProfile,
  ILeaveStatistics
} from '../../../models/ILeaveModels';

/**
 * LeaveAdministration component for managing leave requests
 * Provides functionality for administrators to view, approve, reject, and manage leave requests
 */
const LeaveAdministration: React.FC<ILeaveAdministrationProps> = (props) => {
  // State management
  const [leaveRequests, setLeaveRequests] = useState<ILeaveRequest[]>([]);
  const [filteredRequests, setFilteredRequests] = useState<ILeaveRequest[]>([]);
  const [leaveTypes, setLeaveTypes] = useState<ILeaveType[]>([]);
  const [users, setUsers] = useState<IUserProfile[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string>('');
  const [selectedView, setSelectedView] = useState<string>(props.defaultView || 'pending');
  const [searchText, setSearchText] = useState<string>('');
  const [statusFilter, setStatusFilter] = useState<string>('all');
  const [leaveTypeFilter, setLeaveTypeFilter] = useState<string>('all');
  const [userFilter, setUserFilter] = useState<string>('all');
  const [dateFromFilter, setDateFromFilter] = useState<Date | null>(null);
  const [dateToFilter, setDateToFilter] = useState<Date | null>(null);
  const [selectedRequest, setSelectedRequest] = useState<ILeaveRequest | null>(null);
  const [isPanelOpen, setIsPanelOpen] = useState<boolean>(false);
  const [isDialogOpen, setIsDialogOpen] = useState<boolean>(false);
  const [dialogAction, setDialogAction] = useState<'approve' | 'reject' | 'delete'>('approve');
  const [actionComment, setActionComment] = useState<string>('');
  const [message, setMessage] = useState<{ text: string; type: MessageBarType } | null>(null);
  const [submitting, setSubmitting] = useState<boolean>(false);
  const [statistics, setStatistics] = useState<ILeaveStatistics | null>(null);
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [bulkSelectedItems, setBulkSelectedItems] = useState<ILeaveRequest[]>([]);
  const [isBulkActionDialogOpen, setIsBulkActionDialogOpen] = useState<boolean>(false);
  const [bulkAction, setBulkAction] = useState<'approve' | 'reject'>('approve');

  // Services
  const sharePointService = new SharePointService(props.context);

  // Selection for bulk actions
  const selection = new Selection({
    onSelectionChanged: () => {
      setBulkSelectedItems(selection.getSelection() as ILeaveRequest[]);
    }
  });

  /**
   * Calculate statistics for the dashboard
   */
  const calculateStatistics = (requests: ILeaveRequest[]) => {
    const stats: ILeaveStatistics = {
      totalRequests: requests.length,
      pendingRequests: requests.filter(r => r.approvalStatus === 'Pending').length,
      approvedRequests: requests.filter(r => r.approvalStatus === 'Approved').length,
      rejectedRequests: requests.filter(r => r.approvalStatus === 'Rejected').length,
      totalDaysRequested: requests.reduce((sum, r) => sum + (r.totalDays || 0), 0),
      totalDaysApproved: requests
        .filter(r => r.approvalStatus === 'Approved')
        .reduce((sum, r) => sum + (r.totalDays || 0), 0)
    };

    setStatistics(stats);
  };

  /**
   * Load leave requests from SharePoint
   */
  const loadLeaveRequests = useCallback(async () => {
    try {
      setLoading(true);
      setError('');
      const requests = await sharePointService.getAllLeaveRequests();
      setLeaveRequests(requests);
      calculateStatistics(requests);
    } catch (err) {
      setError('Failed to load leave requests. Please try again.');
      console.error('Error loading leave requests:', err);
    } finally {
      setLoading(false);
    }
  }, [sharePointService]);

  /**
   * Load leave types from SharePoint
   */
  const loadLeaveTypes = useCallback(async () => {
    try {
      const types = await sharePointService.getLeaveTypes();
      setLeaveTypes(types);
    } catch (err) {
      console.error('Error loading leave types:', err);
    }
  }, [sharePointService]);

  /**
   * Load users from SharePoint
   */
  const loadUsers = useCallback(async () => {
    try {
      // This would typically load from User Information List or similar
      // For now, we'll extract unique users from leave requests
      const uniqueUsers = leaveRequests.reduce((acc, request) => {
        let userExists = false;
        for (let i = 0; i < acc.length; i++) {
          if (acc[i].id === request.employeeId) {
            userExists = true;
            break;
          }
        }
        if (!userExists) {
          acc.push({
            id: request.employeeId ? parseInt(request.employeeId, 10) || 0 : 0,
            displayName: request.requesterName,
            email: request.employeeId || '',
            department: request.department || ''
          });
        }
        return acc;
      }, [] as IUserProfile[]);
      setUsers(uniqueUsers);
    } catch (err) {
      console.error('Error loading users:', err);
    }
  }, [leaveRequests]);

  /**
   * Filter requests based on current filters
   */
  const filterRequests = useCallback(() => {
    let filtered = [...leaveRequests];

    // Filter by view
    if (selectedView !== 'all') {
      filtered = filtered.filter(request => {
        switch (selectedView) {
          case 'pending':
            return request.approvalStatus === 'Pending';
          case 'approved':
            return request.approvalStatus === 'Approved';
          case 'rejected':
            return request.approvalStatus === 'Rejected';
          default:
            return true;
        }
      });
    }

    // Filter by search text
    if (searchText) {
      filtered = filtered.filter(request =>
        (request.requesterName && request.requesterName.toLowerCase().indexOf(searchText.toLowerCase()) !== -1) ||
        (request.leaveType && request.leaveType.toLowerCase().indexOf(searchText.toLowerCase()) !== -1) ||
        (request.requestComments && request.requestComments.toLowerCase().indexOf(searchText.toLowerCase()) !== -1)
      );
    }

    // Filter by status
    if (statusFilter !== 'all') {
      filtered = filtered.filter(request => request.approvalStatus === statusFilter);
    }

    // Filter by leave type
    if (leaveTypeFilter !== 'all') {
      filtered = filtered.filter(request => request.leaveType === leaveTypeFilter);
    }

    // Filter by user
    if (userFilter !== 'all') {
      filtered = filtered.filter(request => request.employeeId === userFilter);
    }

    // Filter by date range
    if (dateFromFilter) {
      filtered = filtered.filter(request => request.startDate && new Date(request.startDate) >= dateFromFilter);
    }
    if (dateToFilter) {
      filtered = filtered.filter(request => request.endDate && new Date(request.endDate) <= dateToFilter);
    }

    setFilteredRequests(filtered);
    setCurrentPage(1); // Reset to first page when filtering
  }, [leaveRequests, selectedView, searchText, statusFilter, leaveTypeFilter, userFilter, dateFromFilter, dateToFilter]);

  /**
   * Handle request action (approve/reject/delete)
   */
  const handleRequestAction = async () => {
    if (!selectedRequest) return;

    try {
      setSubmitting(true);
      
      if (dialogAction === 'delete') {
        // Delete request logic would go here
        setMessage({ text: 'Request deleted successfully', type: MessageBarType.success });
      } else {
        await sharePointService.updateLeaveRequestStatus(
          selectedRequest.id || 0,
          dialogAction === 'approve' ? 'Approved' : 'Rejected',
          actionComment
        );
        setMessage({ 
          text: `Request ${dialogAction}d successfully`, 
          type: MessageBarType.success 
        });
      }
      
      await loadLeaveRequests();
      setIsDialogOpen(false);
      setSelectedRequest(null);
      setActionComment('');
    } catch (err) {
      setMessage({ text: 'Failed to update request', type: MessageBarType.error });
      console.error('Error updating request:', err);
    } finally {
      setSubmitting(false);
    }
  };

  /**
   * Handle bulk actions
   */
  const handleBulkAction = async () => {
    if (bulkSelectedItems.length === 0) return;

    try {
      setSubmitting(true);
      
      for (const request of bulkSelectedItems) {
        await sharePointService.updateLeaveRequestStatus(
          request.id || 0,
          bulkAction === 'approve' ? 'Approved' : 'Rejected',
          actionComment
        );
      }
      
      setMessage({ 
        text: `${bulkSelectedItems.length} requests ${bulkAction}d successfully`, 
        type: MessageBarType.success 
      });
      
      await loadLeaveRequests();
      setIsBulkActionDialogOpen(false);
      setBulkSelectedItems([]);
      selection.setAllSelected(false);
      setActionComment('');
    } catch (err) {
      setMessage({ text: 'Failed to update requests', type: MessageBarType.error });
      console.error('Error updating requests:', err);
    } finally {
      setSubmitting(false);
    }
  };

  /**
   * Export data to CSV
   */
  const exportToCSV = () => {
    const headers = ['Employee Name', 'Leave Type', 'Start Date', 'End Date', 'Total Days', 'Status', 'Comments'];
    const csvContent = [
      headers.join(','),
      ...filteredRequests.map(request => [
        request.requesterName || '',
        request.leaveType || '',
        request.startDate ? new Date(request.startDate).toLocaleDateString() : '',
        request.endDate ? new Date(request.endDate).toLocaleDateString() : '',
        request.totalDays || 0,
        request.approvalStatus || '',
        request.requestComments || ''
      ].map(field => `"${field}"`).join(','))
    ].join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `leave-requests-${new Date().toISOString().split('T')[0]}.csv`;
    a.click();
    window.URL.revokeObjectURL(url);
  };

  // Load data on component mount
  useEffect(() => {
    void loadLeaveRequests();
    void loadLeaveTypes();
  }, [loadLeaveRequests, loadLeaveTypes]);

  // Load users when leave requests change
  useEffect(() => {
    if (leaveRequests.length > 0) {
      void loadUsers();
    }
  }, [leaveRequests, loadUsers]);

  // Filter requests when filters change
  useEffect(() => {
    filterRequests();
  }, [filterRequests]);

  // Clear message after 5 seconds
  useEffect(() => {
    if (message) {
      const timer = setTimeout(() => setMessage(null), 5000);
      return () => clearTimeout(timer);
    }
  }, [message]);

  // Command bar items
  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: loadLeaveRequests
    },
    {
      key: 'export',
      text: 'Export',
      iconProps: { iconName: 'Download' },
      onClick: exportToCSV
    }
  ];

  if (props.allowBulkActions && bulkSelectedItems.length > 0) {
    commandBarItems.push(
      {
        key: 'bulkApprove',
        text: `Approve (${bulkSelectedItems.length})`,
        iconProps: { iconName: 'CheckMark' },
        onClick: () => {
          setBulkAction('approve');
          setIsBulkActionDialogOpen(true);
        }
      },
      {
        key: 'bulkReject',
        text: `Reject (${bulkSelectedItems.length})`,
        iconProps: { iconName: 'Cancel' },
        onClick: () => {
          setBulkAction('reject');
          setIsBulkActionDialogOpen(true);
        }
      }
    );
  }

  // Dropdown options
  const statusOptions: IDropdownOption[] = [
    { key: 'all', text: 'All Statuses' },
    { key: 'Pending', text: 'Pending' },
    { key: 'Approved', text: 'Approved' },
    { key: 'Rejected', text: 'Rejected' }
  ];

  const leaveTypeOptions: IDropdownOption[] = [
    { key: 'all', text: 'All Leave Types' },
    ...leaveTypes.map(type => ({ key: type.title || '', text: type.title || '' }))
  ];

  const userOptions: IDropdownOption[] = [
    { key: 'all', text: 'All Users' },
    ...users.map(user => ({ key: user.id || 0, text: user.displayName || '' }))
  ];

  // Table columns
  const columns: IColumn[] = [
    {
      key: 'requesterName',
      name: 'Employee',
      fieldName: 'requesterName',
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      onRender: (item: ILeaveRequest) => (
        <Stack>
          <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
            {item.requesterName}
          </Text>
          {item.department && (
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              {item.department}
            </Text>
          )}
        </Stack>
      )
    },
    {
      key: 'leaveType',
      name: 'Leave Type',
      fieldName: 'leaveType',
      minWidth: 120,
      maxWidth: 150,
      isResizable: true
    },
    {
      key: 'dates',
      name: 'Dates',
      fieldName: 'startDate',
      minWidth: 180,
      maxWidth: 220,
      isResizable: true,
      onRender: (item: ILeaveRequest) => (
            <Stack>
              <Text variant="small">
                {item.startDate ? new Date(item.startDate).toLocaleDateString() : 'N/A'} - {item.endDate ? new Date(item.endDate).toLocaleDateString() : 'N/A'}
              </Text>
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                {item.totalDays || 0} day{(item.totalDays || 0) !== 1 ? 's' : ''}
              </Text>
            </Stack>
          )
    },
    {
      key: 'status',
      name: 'Status',
      fieldName: 'approvalStatus',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
      onRender: (item: ILeaveRequest) => {
        const getStatusColor = (status: string) => {
          switch (status) {
            case 'Approved': return '#107c10';
            case 'Rejected': return '#d13438';
            case 'Pending': return '#ff8c00';
            default: return '#605e5c';
          }
        };

        const getStatusIcon = (status: string) => {
          switch (status) {
            case 'Approved': return 'CheckMark';
            case 'Rejected': return 'Cancel';
            case 'Pending': return 'Clock';
            default: return 'Info';
          }
        };

        return (
          <div className={styles.statusCell}>
            <Icon 
              iconName={getStatusIcon(item.approvalStatus || '')} 
              styles={{ root: { color: getStatusColor(item.approvalStatus || ''), marginRight: 8 } }} 
            />
            <span style={{ color: getStatusColor(item.approvalStatus || ''), fontWeight: 500 }}>
              {item.approvalStatus || 'Unknown'}
            </span>
          </div>
        );
      }
    },
    {
      key: 'submissionDate',
      name: 'Submitted',
      fieldName: 'submissionDate',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
      onRender: (item: ILeaveRequest) => (
        <Text variant="small">
          {item.submissionDate ? new Date(item.submissionDate).toLocaleDateString() : 'N/A'}
        </Text>
      )
    },
    {
      key: 'actions',
      name: 'Actions',
      fieldName: '',
      minWidth: 120,
      maxWidth: 150,
      isResizable: false,
      onRender: (item: ILeaveRequest) => (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <TooltipHost content="View Details">
            <Icon 
              iconName="View" 
              styles={{ root: { cursor: 'pointer', color: '#0078d4' } }}
              onClick={() => {
                setSelectedRequest(item);
                setIsPanelOpen(true);
              }}
            />
          </TooltipHost>
          {item.approvalStatus === 'Pending' && (
            <>
              <TooltipHost content="Approve">
                <Icon 
                  iconName="CheckMark" 
                  styles={{ root: { cursor: 'pointer', color: '#107c10' } }}
                  onClick={() => {
                    setSelectedRequest(item);
                    setDialogAction('approve');
                    setIsDialogOpen(true);
                  }}
                />
              </TooltipHost>
              <TooltipHost content="Reject">
                <Icon 
                  iconName="Cancel" 
                  styles={{ root: { cursor: 'pointer', color: '#d13438' } }}
                  onClick={() => {
                    setSelectedRequest(item);
                    setDialogAction('reject');
                    setIsDialogOpen(true);
                  }}
                />
              </TooltipHost>
            </>
          )}
        </Stack>
      )
    }
  ];

  // Pagination
  const itemsPerPage = props.itemsPerPage || 10;
  const totalPages = Math.ceil(filteredRequests.length / itemsPerPage);
  const startIndex = (currentPage - 1) * itemsPerPage;
  const endIndex = startIndex + itemsPerPage;
  const currentItems = filteredRequests.slice(startIndex, endIndex);



  if (loading) {
    return (
      <div className={styles.loadingContainer}>
        <Spinner size={SpinnerSize.large} label="Loading leave requests..." />
      </div>
    );
  }

  return (
    <div className={`${styles.leaveAdministration} ${props.isDarkTheme ? styles.dark : ''}`}>
      {props.title && (
        <Text variant="xxLarge" className={styles.title}>
          {props.title}
        </Text>
      )}

      {message && (
        <MessageBar
          messageBarType={message.type}
          onDismiss={() => setMessage(null)}
          styles={{ root: { marginBottom: 16 } }}
        >
          {message.text}
        </MessageBar>
      )}

      {error && (
        <div className={styles.errorContainer}>
          <MessageBar messageBarType={MessageBarType.error}>
            {error}
          </MessageBar>
        </div>
      )}

      <div className={styles.pivotContainer}>
        <Pivot
          selectedKey={selectedView}
          onLinkClick={(item) => setSelectedView(item?.props.itemKey || 'all')}
        >
          <PivotItem headerText="Dashboard" itemKey="dashboard" />
          <PivotItem headerText="All Requests" itemKey="all" />
          <PivotItem headerText="Pending" itemKey="pending" />
          <PivotItem headerText="Approved" itemKey="approved" />
          <PivotItem headerText="Rejected" itemKey="rejected" />
        </Pivot>
      </div>

      {selectedView === 'dashboard' && statistics ? (
        <div className={styles.analyticsContainer}>
          <div className={styles.statsGrid}>
            <div className={styles.statCard}>
              <Text variant="xxLarge" className={styles.statNumber}>
                {statistics.totalRequests}
              </Text>
              <Text variant="medium">Total Requests</Text>
            </div>
            <div className={styles.statCard}>
              <Text variant="xxLarge" className={styles.statNumber}>
                {statistics.pendingRequests}
              </Text>
              <Text variant="medium">Pending</Text>
            </div>
            <div className={styles.statCard}>
              <Text variant="xxLarge" className={styles.statNumber}>
                {statistics.approvedRequests}
              </Text>
              <Text variant="medium">Approved</Text>
            </div>
            <div className={styles.statCard}>
              <Text variant="xxLarge" className={styles.statNumber}>
                {statistics.rejectedRequests}
              </Text>
              <Text variant="medium">Rejected</Text>
            </div>
            <div className={styles.statCard}>
              <Text variant="xxLarge" className={styles.statNumber}>
                {statistics.totalDaysRequested}
              </Text>
              <Text variant="medium">Total Days</Text>
            </div>
          </div>

          <Stack horizontal tokens={{ childrenGap: 20 }} wrap>
            <div className={styles.chartCard}>
              <Text variant="large" styles={{ root: { fontWeight: 600, marginBottom: 16 } }}>
                Requests by Status
              </Text>
              <ProgressIndicator 
                label="Approved" 
                percentComplete={((statistics.approvedRequests || 0) / (statistics.totalRequests || 1))}
                styles={{ root: { marginBottom: 8 } }}
              />
              <ProgressIndicator 
                label="Pending" 
                percentComplete={((statistics.pendingRequests || 0) / (statistics.totalRequests || 1))}
                styles={{ root: { marginBottom: 8 } }}
              />
              <ProgressIndicator 
                label="Rejected" 
                percentComplete={((statistics.rejectedRequests || 0) / (statistics.totalRequests || 1))}
              />
            </div>

            <div className={styles.chartCard}>
              <Text variant="large" styles={{ root: { fontWeight: 600, marginBottom: 16 } }}>
                Leave Types
              </Text>
              <Stack horizontal horizontalAlign="space-between" styles={{ root: { marginBottom: 8 } }}>
                <Text>Total Requests</Text>
                <Text styles={{ root: { fontWeight: 600 } }}>{statistics.totalRequests || 0}</Text>
              </Stack>
              <Stack horizontal horizontalAlign="space-between" styles={{ root: { marginBottom: 8 } }}>
                <Text>Days Approved</Text>
                <Text styles={{ root: { fontWeight: 600 } }}>{statistics.totalDaysApproved || 0}</Text>
              </Stack>
            </div>
          </Stack>
        </div>
      ) : (
        <>
          <div className={styles.commandBarContainer}>
            <CommandBar items={commandBarItems} />
          </div>

          <div className={styles.filterContainer}>
            <div className={styles.filterRow}>
              <SearchBox
                placeholder="Search by employee name, leave type, or comments"
                value={searchText}
                onChange={(_, newValue) => setSearchText(newValue || '')}
                styles={{ root: { width: 300 } }}
              />
              <Dropdown
                label="Status"
                options={statusOptions}
                selectedKey={statusFilter}
                onChange={(_, option) => setStatusFilter(option?.key as string)}
                styles={{ root: { width: 150 } }}
              />
              <Dropdown
                label="Leave Type"
                options={leaveTypeOptions}
                selectedKey={leaveTypeFilter}
                onChange={(_, option) => setLeaveTypeFilter(option?.key as string)}
                styles={{ root: { width: 150 } }}
              />
              <Dropdown
                label="Employee"
                options={userOptions}
                selectedKey={userFilter}
                onChange={(_, option) => setUserFilter(option?.key as string)}
                styles={{ root: { width: 200 } }}
              />
              <DatePicker
                label="From Date"
                value={dateFromFilter || undefined}
                onSelectDate={(date) => setDateFromFilter(date || null)}
                styles={{ root: { width: 150 } }}
              />
              <DatePicker
                label="To Date"
                value={dateToFilter || undefined}
                onSelectDate={(date) => setDateToFilter(date || null)}
                styles={{ root: { width: 150 } }}
              />
            </div>
          </div>

          <div className={styles.listContainer}>
            {filteredRequests.length === 0 ? (
              <div className={styles.emptyState}>
                <Icon iconName="Calendar" className={styles.emptyIcon} />
                <Text variant="large" className={styles.emptyTitle}>
                  No leave requests found
                </Text>
                <Text className={styles.emptyDescription}>
                  Try adjusting your filters or check back later.
                </Text>
              </div>
            ) : (
              <>
                <DetailsList
                  items={currentItems}
                  columns={columns}
                  selectionMode={props.allowBulkActions ? SelectionMode.multiple : SelectionMode.none}
                  selection={props.allowBulkActions ? selection : undefined}
                  setKey="set"
                  layoutMode={0}
                  isHeaderVisible={true}
                />

                {totalPages > 1 && (
                  <div className={styles.paginationContainer}>
                    <DefaultButton
                      text="Previous"
                      disabled={currentPage === 1}
                      onClick={() => setCurrentPage(currentPage - 1)}
                    />
                    <Text>
                      Page {currentPage} of {totalPages} ({filteredRequests.length} total items)
                    </Text>
                    <DefaultButton
                      text="Next"
                      disabled={currentPage === totalPages}
                      onClick={() => setCurrentPage(currentPage + 1)}
                    />
                  </div>
                )}
              </>
            )}
          </div>
        </>
      )}

      {/* Details Panel */}
      <Panel
        isOpen={isPanelOpen}
        onDismiss={() => setIsPanelOpen(false)}
        type={PanelType.medium}
        headerText="Leave Request Details"
        closeButtonAriaLabel="Close"
      >
        {selectedRequest && (
          <div className={styles.detailsPanel}>
            <div className={styles.detailRow}>
              <Label>Employee</Label>
              <Text>{selectedRequest.requesterName}</Text>
            </div>
            <div className={styles.detailRow}>
              <Label>Leave Type</Label>
              <Text>{selectedRequest.leaveType}</Text>
            </div>
            <div className={styles.detailRow}>
              <Label>Start Date</Label>
              <Text>{selectedRequest.startDate ? new Date(selectedRequest.startDate).toLocaleDateString() : 'N/A'}</Text>
            </div>
            <div className={styles.detailRow}>
              <Label>End Date</Label>
              <Text>{selectedRequest.endDate ? new Date(selectedRequest.endDate).toLocaleDateString() : 'N/A'}</Text>
            </div>
            <div className={styles.detailRow}>
              <Label>Total Days</Label>
              <Text>{selectedRequest.totalDays}</Text>
            </div>
            <div className={styles.detailRow}>
              <Label>Status</Label>
              <div className={`${styles.statusBadge} ${selectedRequest.approvalStatus ? styles[selectedRequest.approvalStatus.toLowerCase() as keyof typeof styles] || '' : ''}`}>
                {selectedRequest.approvalStatus}
              </div>
            </div>
            {selectedRequest.requestComments && (
              <div className={styles.detailRow}>
                <Label>Comments</Label>
                <Text>{selectedRequest.requestComments}</Text>
              </div>
            )}
            <div className={styles.detailRow}>
              <Label>Submitted Date</Label>
              <Text>{selectedRequest.submissionDate ? new Date(selectedRequest.submissionDate).toLocaleDateString() : 'N/A'}</Text>
            </div>
            
            {selectedRequest.approvalStatus === 'Pending' && (
              <Stack horizontal tokens={{ childrenGap: 10 }} styles={{ root: { marginTop: 20 } }}>
                <PrimaryButton
                  text="Approve"
                  onClick={() => {
                    setDialogAction('approve');
                    setIsDialogOpen(true);
                    setIsPanelOpen(false);
                  }}
                />
                <DefaultButton
                  text="Reject"
                  onClick={() => {
                    setDialogAction('reject');
                    setIsDialogOpen(true);
                    setIsPanelOpen(false);
                  }}
                />
              </Stack>
            )}
          </div>
        )}
      </Panel>

      {/* Action Dialog */}
      <Dialog
        hidden={!isDialogOpen}
        onDismiss={() => setIsDialogOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: `${dialogAction.charAt(0).toUpperCase() + dialogAction.slice(1)} Leave Request`,
          subText: selectedRequest ? 
            `Are you sure you want to ${dialogAction} the leave request for ${selectedRequest.requesterName}?` :
            ''
        }}
      >
        <TextField
          label="Comments (optional)"
          multiline
          rows={3}
          value={actionComment}
          onChange={(_, newValue) => setActionComment(newValue || '')}
          placeholder={`Add a comment for this ${dialogAction}...`}
        />
        <DialogFooter>
          <PrimaryButton
            text={dialogAction.charAt(0).toUpperCase() + dialogAction.slice(1)}
            onClick={handleRequestAction}
            disabled={submitting}
          />
          <DefaultButton
            text="Cancel"
            onClick={() => setIsDialogOpen(false)}
            disabled={submitting}
          />
        </DialogFooter>
      </Dialog>

      {/* Bulk Action Dialog */}
      <Dialog
        hidden={!isBulkActionDialogOpen}
        onDismiss={() => setIsBulkActionDialogOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: `${bulkAction.charAt(0).toUpperCase() + bulkAction.slice(1)} Multiple Requests`,
          subText: `Are you sure you want to ${bulkAction} ${bulkSelectedItems.length} selected requests?`
        }}
      >
        <TextField
          label="Comments (optional)"
          multiline
          rows={3}
          value={actionComment}
          onChange={(_, newValue) => setActionComment(newValue || '')}
          placeholder={`Add a comment for this bulk ${bulkAction}...`}
        />
        <DialogFooter>
          <PrimaryButton
            text={`${bulkAction.charAt(0).toUpperCase() + bulkAction.slice(1)} All`}
            onClick={handleBulkAction}
            disabled={submitting}
          />
          <DefaultButton
            text="Cancel"
            onClick={() => setIsBulkActionDialogOpen(false)}
            disabled={submitting}
          />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default LeaveAdministration;