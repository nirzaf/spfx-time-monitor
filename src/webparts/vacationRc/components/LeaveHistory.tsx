import * as React from 'react';
import { useState } from 'react';
import {
  Stack,
  Text,
  CommandBar,
  ICommandBarItemProps,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  Dropdown,
  IDropdownOption,
  TextField,
  DatePicker,
  Panel,
  PanelType,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Pivot,
  PivotItem,
  SearchBox,
  Icon
} from '@fluentui/react';

import type { ILeaveHistoryProps } from './ILeaveHistoryProps';
import { SharePointService } from '../../../services/SharePointService';
import { ILeaveRequest, ILeaveStatistics } from '../../../models/ILeaveModels';
import * as moment from 'moment';



interface ILeaveHistoryState {
  leaveRequests: ILeaveRequest[];
  filteredRequests: ILeaveRequest[];
  loading: boolean;
  error: string | null;
  selectedView: string;
  searchText: string;
  selectedStatus: string;
  selectedLeaveType: string;
  selectedYear: string;
  startDate: Date | null;
  endDate: Date | null;
  selectedRequest: ILeaveRequest | null;
  showDetailsPanel: boolean;
  statistics: ILeaveStatistics | null;
  currentPage: number;
}

const LeaveHistory: React.FC<ILeaveHistoryProps> = (props) => {
  const {
    context,
    title,
    defaultView,
    itemsPerPage,
    showAnalytics
  } = props;

  const [state, setState] = useState<ILeaveHistoryState>({
    leaveRequests: [],
    filteredRequests: [],
    loading: true,
    error: null,
    selectedView: defaultView || 'list',
    searchText: '',
    selectedStatus: 'All',
    selectedLeaveType: 'All',
    selectedYear: new Date().getFullYear().toString(),
    startDate: null,
    endDate: null,
    selectedRequest: null,
    showDetailsPanel: false,
    statistics: null,
    currentPage: 1
  });

  const sharePointService = new SharePointService(context);

  React.useEffect(() => {
    loadLeaveHistory();
  }, []);

  React.useEffect(() => {
    applyFilters();
  }, [state.searchText, state.selectedStatus, state.selectedLeaveType, state.selectedYear, state.startDate, state.endDate, state.leaveRequests]);

  React.useEffect(() => {
    calculateStatistics();
  }, [state.filteredRequests]);

  const loadLeaveHistory = async (): Promise<void> => {
    try {
      setState(prev => ({ ...prev, loading: true, error: null }));
      
      // Get current user ID from context
      const currentUser = await context.msGraphClientFactory.getClient().then(client => 
        client.api('/me').get()
      );
      const userId = currentUser.id;
      
      const requests = await sharePointService.getUserLeaveRequests(parseInt(userId));
      
      setState(prev => ({
        ...prev,
        leaveRequests: requests,
        loading: false
      }));
    } catch (error) {
      setState(prev => ({
        ...prev,
        error: 'Failed to load leave history',
        loading: false
      }));
    }
  };

  const applyFilters = (): void => {
    let filtered = [...state.leaveRequests];

    // Text search
    if (state.searchText) {
      const searchLower = state.searchText.toLowerCase();
      filtered = filtered.filter(req => 
        req.leaveType?.toLowerCase().indexOf(searchLower) !== -1 ||
        req.requestComments?.toLowerCase().indexOf(searchLower) !== -1 ||
        req.approvalStatus?.toLowerCase().indexOf(searchLower) !== -1
      );
    }

    // Status filter
    if (state.selectedStatus !== 'All') {
      filtered = filtered.filter(req => req.approvalStatus === state.selectedStatus);
    }

    // Leave type filter
    if (state.selectedLeaveType !== 'All') {
      filtered = filtered.filter(req => req.leaveType === state.selectedLeaveType);
    }

    // Year filter
    if (state.selectedYear !== 'All') {
      filtered = filtered.filter(req => 
        moment(req.startDate).year().toString() === state.selectedYear
      );
    }

    // Date range filter
    if (state.startDate) {
      filtered = filtered.filter(req => 
        moment(req.startDate).isSameOrAfter(moment(state.startDate))
      );
    }

    if (state.endDate) {
      filtered = filtered.filter(req => 
        moment(req.endDate).isSameOrBefore(moment(state.endDate))
      );
    }

    // Sort by most recent first
    filtered.sort((a, b) => moment(b.submissionDate).diff(moment(a.submissionDate)));

    setState(prev => ({ ...prev, filteredRequests: filtered, currentPage: 1 }));
  };

  const calculateStatistics = (): void => {
    const requests = state.filteredRequests;
    
    const totalRequests = requests.length;
    const approvedRequests = requests.filter(r => r.approvalStatus === 'Approved').length;
    const pendingRequests = requests.filter(r => r.approvalStatus === 'Pending').length;
    const rejectedRequests = requests.filter(r => r.approvalStatus === 'Rejected').length;
    
    const totalDaysRequested = requests
      .filter(r => r.approvalStatus === 'Approved')
      .reduce((sum, r) => sum + (r.totalDays || 0), 0);

    const statistics: ILeaveStatistics = {
      totalRequests,
      approvedRequests,
      pendingRequests,
      rejectedRequests,
      totalDaysRequested,
      totalDaysApproved: totalDaysRequested,
      averageProcessingTime: 0 // Placeholder since we don't have processing time data
    };

    setState(prev => ({ ...prev, statistics }));
  };

  const exportToCSV = (): void => {
    const csvContent = [
      ['Leave Type', 'Start Date', 'End Date', 'Days', 'Status', 'Submitted Date', 'Comments'],
      ...state.filteredRequests.map(req => [
        req.leaveType || '',
        req.startDate || '',
        req.endDate || '',
        req.totalDays?.toString() || '0',
        req.approvalStatus || '',
        moment(req.submissionDate).format('YYYY-MM-DD'),
        req.requestComments || ''
      ])
    ].map(row => row.join(',')).join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `leave-history-${moment().format('YYYY-MM-DD')}.csv`;
    a.click();
    window.URL.revokeObjectURL(url);
  };

  const getStatusIcon = (status: string): string => {
    switch (status) {
      case 'Approved': return 'CheckMark';
      case 'Rejected': return 'Cancel';
      case 'Pending': return 'Clock';
      default: return 'Info';
    }
  };

  const getStatusColor = (status: string): string => {
    switch (status) {
      case 'Approved': return '#107c10';
      case 'Rejected': return '#d13438';
      case 'Pending': return '#ff8c00';
      default: return '#666666';
    }
  };

  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: () => loadLeaveHistory()
    },
    {
      key: 'export',
      text: 'Export',
      iconProps: { iconName: 'Download' },
      onClick: exportToCSV
    }
  ];

  const statusOptions: IDropdownOption[] = [
    { key: 'All', text: 'All Statuses' },
    { key: 'Pending', text: 'Pending' },
    { key: 'Approved', text: 'Approved' },
    { key: 'Rejected', text: 'Rejected' }
  ];

  const leaveTypeOptions: IDropdownOption[] = [
    { key: 'All', text: 'All Leave Types' },
    { key: 'Annual Leave', text: 'Annual Leave' },
    { key: 'Sick Leave', text: 'Sick Leave' },
    { key: 'Personal Leave', text: 'Personal Leave' },
    { key: 'Maternity/Paternity', text: 'Maternity/Paternity' },
    { key: 'Emergency Leave', text: 'Emergency Leave' }
  ];

  const yearOptions: IDropdownOption[] = [
    { key: 'All', text: 'All Years' },
    { key: '2024', text: '2024' },
    { key: '2023', text: '2023' },
    { key: '2022', text: '2022' },
    { key: '2021', text: '2021' },
    { key: '2020', text: '2020' }
  ];

  const listColumns: IColumn[] = [
    {
      key: 'status',
      name: 'Status',
      fieldName: 'approvalStatus',
      minWidth: 80,
      maxWidth: 100,
      onRender: (item: ILeaveRequest) => (
        <div style={{ display: 'flex', alignItems: 'center' }}>
          <Icon 
            iconName={getStatusIcon(item.approvalStatus || '')} 
            style={{ color: getStatusColor(item.approvalStatus || ''), marginRight: 8 }}
          />
          <span>{item.approvalStatus}</span>
        </div>
      )
    },
    {
      key: 'leaveType',
      name: 'Leave Type',
      fieldName: 'leaveType',
      minWidth: 120,
      maxWidth: 150
    },
    {
      key: 'startDate',
      name: 'Start Date',
      fieldName: 'startDate',
      minWidth: 100,
      maxWidth: 120,
      onRender: (item: ILeaveRequest) => moment(item.startDate).format('MMM DD, YYYY')
    },
    {
      key: 'endDate',
      name: 'End Date',
      fieldName: 'endDate',
      minWidth: 100,
      maxWidth: 120,
      onRender: (item: ILeaveRequest) => moment(item.endDate).format('MMM DD, YYYY')
    },
    {
      key: 'days',
      name: 'Days',
      fieldName: 'totalDays',
      minWidth: 60,
      maxWidth: 80
    },
    {
      key: 'submitted',
      name: 'Submitted',
      fieldName: 'submissionDate',
      minWidth: 100,
      maxWidth: 120,
      onRender: (item: ILeaveRequest) => moment(item.submissionDate).format('MMM DD, YYYY')
    },
    {
      key: 'actions',
      name: 'Actions',
      minWidth: 80,
      maxWidth: 100,
      onRender: (item: ILeaveRequest) => (
        <DefaultButton
          text="View"
          onClick={() => setState(prev => ({ 
            ...prev, 
            selectedRequest: item, 
            showDetailsPanel: true 
          }))}
        />
      )
    }
  ];



  // Pagination
  const totalPages = Math.ceil(state.filteredRequests.length / itemsPerPage);
  const startIndex = (state.currentPage - 1) * itemsPerPage;
  const endIndex = startIndex + itemsPerPage;
  const currentItems = state.filteredRequests.slice(startIndex, endIndex);



  if (state.loading) {
    return (
      <div style={{ padding: '20px' }}>
        <Stack horizontalAlign="center" verticalAlign="center" style={{ height: '200px' }}>
          <Spinner size={SpinnerSize.large} label="Loading leave history..." />
        </Stack>
      </div>
    );
  }

  return (
    <div style={{ padding: '20px' }}>
      <Stack tokens={{ childrenGap: 20 }}>
        {title && (
          <Text variant="xxLarge" style={{ fontWeight: 'bold', marginBottom: '10px' }}>
            {title}
          </Text>
        )}
        
        {state.error && (
          <MessageBar messageBarType={MessageBarType.error}>
            {state.error}
          </MessageBar>
        )}

        <CommandBar items={commandBarItems} />

        {/* Filters */}
        <div style={{ padding: '16px', border: '1px solid #e1e1e1', borderRadius: '4px' }}>
          <Stack tokens={{ childrenGap: 15 }}>
            <Text variant="large">Filters</Text>
            <Stack horizontal tokens={{ childrenGap: 15 }} wrap>
              <SearchBox
                placeholder="Search leave requests..."
                value={state.searchText}
                onChange={(_, newValue) => setState(prev => ({ ...prev, searchText: newValue || '' }))}
                styles={{ root: { width: 200 } }}
              />
              <Dropdown
                placeholder="Status"
                options={statusOptions}
                selectedKey={state.selectedStatus}
                onChange={(_, option) => setState(prev => ({ ...prev, selectedStatus: option?.key as string }))}
                styles={{ root: { width: 150 } }}
              />
              <Dropdown
                placeholder="Leave Type"
                options={leaveTypeOptions}
                selectedKey={state.selectedLeaveType}
                onChange={(_, option) => setState(prev => ({ ...prev, selectedLeaveType: option?.key as string }))}
                styles={{ root: { width: 150 } }}
              />
              <Dropdown
                placeholder="Year"
                options={yearOptions}
                selectedKey={state.selectedYear}
                onChange={(_, option) => setState(prev => ({ ...prev, selectedYear: option?.key as string }))}
                styles={{ root: { width: 100 } }}
              />
              <DatePicker
                placeholder="Start Date"
                value={state.startDate || undefined}
                onSelectDate={(date) => setState(prev => ({ ...prev, startDate: date || null }))}
                styles={{ root: { width: 120 } }}
              />
              <DatePicker
                placeholder="End Date"
                value={state.endDate || undefined}
                onSelectDate={(date) => setState(prev => ({ ...prev, endDate: date || null }))}
                styles={{ root: { width: 120 } }}
              />
              <DefaultButton
                text="Clear Filters"
                onClick={() => setState(prev => ({
                  ...prev,
                  searchText: '',
                  selectedStatus: 'All',
                  selectedLeaveType: 'All',
                  selectedYear: 'All',
                  startDate: null,
                  endDate: null
                }))}
              />
            </Stack>
          </Stack>
        </div>

        <Pivot
          selectedKey={state.selectedView}
          onLinkClick={(item) => setState(prev => ({ ...prev, selectedView: item?.props.itemKey || 'list' }))}
        >
          <PivotItem headerText="List View" itemKey="list">
            <Stack tokens={{ childrenGap: 15 }}>
              <Text variant="medium">
                Showing {currentItems.length} of {state.filteredRequests.length} requests
              </Text>
              
              <DetailsList
                items={currentItems}
                columns={listColumns}
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
                compact={false}
              />

              {/* Pagination */}
              {totalPages > 1 && (
                <Stack horizontal horizontalAlign="center" tokens={{ childrenGap: 10 }}>
                  <DefaultButton
                    text="Previous"
                    disabled={state.currentPage === 1}
                    onClick={() => setState(prev => ({ ...prev, currentPage: prev.currentPage - 1 }))}
                  />
                  <Text>Page {state.currentPage} of {totalPages}</Text>
                  <DefaultButton
                    text="Next"
                    disabled={state.currentPage === totalPages}
                    onClick={() => setState(prev => ({ ...prev, currentPage: prev.currentPage + 1 }))}
                  />
                </Stack>
              )}
            </Stack>
          </PivotItem>

          {showAnalytics && (
            <PivotItem headerText="Analytics" itemKey="analytics">
              <Stack tokens={{ childrenGap: 20 }}>
                {/* Statistics Cards */}
                <Stack horizontal tokens={{ childrenGap: 20 }} wrap>
                  <div style={{ padding: '16px', border: '1px solid #e1e1e1', borderRadius: '4px' }}>
                    <Stack horizontalAlign="center">
                      <Text variant="xxLarge" style={{ fontWeight: 'bold', fontSize: '32px' }}>
                        {state.statistics?.totalRequests || 0}
                      </Text>
                      <Text variant="medium">Total Requests</Text>
                    </Stack>
                  </div>
                  <div style={{ padding: '16px', border: '1px solid #e1e1e1', borderRadius: '4px' }}>
                    <Stack horizontalAlign="center">
                      <Text variant="xxLarge" style={{ color: '#107c10', fontWeight: 'bold', fontSize: '32px' }}>
                        {state.statistics?.approvedRequests || 0}
                      </Text>
                      <Text variant="medium">Approved</Text>
                    </Stack>
                  </div>
                  <div style={{ padding: '16px', border: '1px solid #e1e1e1', borderRadius: '4px' }}>
                    <Stack horizontalAlign="center">
                      <Text variant="xxLarge" style={{ color: '#ff8c00', fontWeight: 'bold', fontSize: '32px' }}>
                        {state.statistics?.pendingRequests || 0}
                      </Text>
                      <Text variant="medium">Pending</Text>
                    </Stack>
                  </div>
                  <div style={{ padding: '16px', border: '1px solid #e1e1e1', borderRadius: '4px' }}>
                    <Stack horizontalAlign="center">
                      <Text variant="xxLarge" style={{ fontWeight: 'bold', fontSize: '32px' }}>
                        {state.statistics?.totalDaysRequested || 0}
                      </Text>
                      <Text variant="medium">Days Requested</Text>
                    </Stack>
                  </div>
                </Stack>

                {/* Charts */}
                <Stack horizontal tokens={{ childrenGap: 20 }} wrap>
                  <div style={{ padding: '16px', border: '1px solid #e1e1e1', borderRadius: '4px' }}>
                    <Stack>
                      <Text variant="large">Leave Types Breakdown</Text>
                      <div style={{ padding: '20px', textAlign: 'center' }}>
                        {/* Chart component would go here */}
                        <Text>Chart visualization coming soon</Text>
                      </div>
                    </Stack>
                  </div>
                  <div style={{ padding: '16px', border: '1px solid #e1e1e1', borderRadius: '4px' }}>
                    <Stack>
                      <Text variant="large">Monthly Usage</Text>
                      <div style={{ padding: '20px', textAlign: 'center' }}>
                        {/* Chart component would go here */}
                        <Text>Chart visualization coming soon</Text>
                      </div>
                    </Stack>
                  </div>
                </Stack>
              </Stack>
            </PivotItem>
          )}
        </Pivot>

        {/* Details Panel */}
        <Panel
          isOpen={state.showDetailsPanel}
          onDismiss={() => setState(prev => ({ ...prev, showDetailsPanel: false }))}
          type={PanelType.medium}
          headerText="Leave Request Details"
        >
          {state.selectedRequest && (
            <Stack tokens={{ childrenGap: 15 }}>
              <TextField label="Leave Type" value={state.selectedRequest.leaveType} readOnly />
              <TextField label="Start Date" value={moment(state.selectedRequest.startDate).format('MMM DD, YYYY')} readOnly />
              <TextField label="End Date" value={moment(state.selectedRequest.endDate).format('MMM DD, YYYY')} readOnly />
              <TextField label="Total Days" value={state.selectedRequest.totalDays?.toString() || '0'} readOnly />
              <TextField label="Status" value={state.selectedRequest.approvalStatus} readOnly />
              <TextField label="Submitted Date" value={moment(state.selectedRequest.submissionDate).format('MMM DD, YYYY')} readOnly />
              {state.selectedRequest.requestComments && (
                <TextField label="Comments" value={state.selectedRequest.requestComments} multiline rows={3} readOnly />
              )}
              {state.selectedRequest.approvalComments && (
                <TextField label="Manager Comments" value={state.selectedRequest.approvalComments} multiline rows={3} readOnly />
              )}
              <DefaultButton
                text="Close"
                onClick={() => setState(prev => ({ ...prev, showDetailsPanel: false }))}
              />
            </Stack>
          )}
        </Panel>
      </Stack>
    </div>
  );
};

export default LeaveHistory;