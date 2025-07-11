import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import FullCalendar from '@fullcalendar/react';
import dayGridPlugin from '@fullcalendar/daygrid';
import timeGridPlugin from '@fullcalendar/timegrid';
import interactionPlugin from '@fullcalendar/interaction';
import { EventInput, EventClickArg, DateSelectArg } from '@fullcalendar/core';
import {
  Stack,
  Text,
  CommandBar,
  ICommandBarItemProps,
  Panel,
  PanelType,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  Dropdown,
  IDropdownOption,
  TextField
} from '@fluentui/react';
import styles from './TeamCalendar.module.scss';
import type { ITeamCalendarProps } from './ITeamCalendarProps';
import { SharePointService } from '../../../services/SharePointService';
import { ILeaveRequest, IUserProfile } from '../../../models/ILeaveModels';
import * as moment from 'moment';

interface ITeamCalendarState {
  events: EventInput[];
  loading: boolean;
  error: string | null;
  selectedEvent: ILeaveRequest | null;
  showEventPanel: boolean;
  showFiltersPanel: boolean;
  leaveRequests: ILeaveRequest[];
  filteredRequests: ILeaveRequest[];
  selectedDepartment: string;
  selectedStatus: string;
  searchText: string;
  teamMembers: IUserProfile[];
}

const TeamCalendar: React.FC<ITeamCalendarProps> = (props) => {
  const {
    context,
    title,
    defaultView,
    showWeekends,
    allowExport,
    isDarkTheme
  } = props;

  const [state, setState] = useState<ITeamCalendarState>({
    events: [],
    loading: true,
    error: null,
    selectedEvent: null,
    showEventPanel: false,
    showFiltersPanel: false,
    leaveRequests: [],
    filteredRequests: [],
    selectedDepartment: 'All',
    selectedStatus: 'All',
    searchText: '',
    teamMembers: []
  });

  const calendarRef = useRef<FullCalendar>(null);
  const sharePointService = new SharePointService(context);

  useEffect(() => {
    loadCalendarData();
  }, []);

  useEffect(() => {
    filterRequests();
  }, [state.leaveRequests, state.selectedDepartment, state.selectedStatus, state.searchText]);

  const loadCalendarData = async (): Promise<void> => {
    try {
      setState(prev => ({ ...prev, loading: true, error: null }));
      
      const requests = await sharePointService.getAllLeaveRequests();
      const events = requests
        .filter(req => req.approvalStatus === 'Approved')
      .map(req => ({
        id: (req.id || 0).toString(),
        title: `${req.requesterName || 'Unknown'} - ${req.leaveType || 'Unknown'}`,
        start: req.startDate,
        end: moment(req.endDate).add(1, 'day').format('YYYY-MM-DD'),
        backgroundColor: getEventColor(req.leaveType || ''),
        borderColor: getEventColor(req.leaveType || ''),
          extendedProps: {
            leaveRequest: req
          }
        }));

      setState(prev => ({
        ...prev,
        events,
        leaveRequests: requests,
        loading: false
      }));
    } catch (error) {
      setState(prev => ({
        ...prev,
        error: 'Failed to load calendar data',
        loading: false
      }));
    }
  };

  const getEventColor = (leaveType: string): string => {
    const colors: { [key: string]: string } = {
      'Annual Leave': '#0078d4',
      'Sick Leave': '#d13438',
      'Personal Leave': '#107c10',
      'Maternity/Paternity': '#8764b8',
      'Emergency Leave': '#ff8c00'
    };
    return colors[leaveType] || '#666666';
  };

  const filterRequests = (): void => {
    let filtered = [...state.leaveRequests];

    if (state.selectedDepartment !== 'All') {
      filtered = filtered.filter(req => req.department === state.selectedDepartment);
    }

    if (state.selectedStatus !== 'All') {
      filtered = filtered.filter(req => req.approvalStatus === state.selectedStatus);
    }

    if (state.searchText) {
      const searchLower = state.searchText.toLowerCase();
      filtered = filtered.filter(req =>
        (req.requesterName && req.requesterName.toLowerCase().indexOf(searchLower) !== -1) ||
        (req.leaveType && req.leaveType.toLowerCase().indexOf(searchLower) !== -1)
      );
    }

    setState(prev => ({ ...prev, filteredRequests: filtered }));
  };

  const handleEventClick = (clickInfo: EventClickArg): void => {
    const leaveRequest = clickInfo.event.extendedProps.leaveRequest as ILeaveRequest;
    setState(prev => ({
      ...prev,
      selectedEvent: leaveRequest,
      showEventPanel: true
    }));
  };

  const handleDateSelect = (selectInfo: DateSelectArg): void => {
    // Handle date selection for creating new events
    console.log('Date selected:', selectInfo);
  };

  const exportToCSV = (): void => {
    const csvContent = [
      ['Employee', 'Leave Type', 'Start Date', 'End Date', 'Status', 'Days'].join(','),
      ...state.filteredRequests.map(req => [
        req.requesterName || '',
        req.leaveType || '',
        req.startDate || '',
        req.endDate || '',
        req.approvalStatus || '',
        (req.totalDays || 0).toString()
      ].map(field => `"${field}"`).join(','))
    ].join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `team-calendar-${moment().format('YYYY-MM-DD')}.csv`;
    a.click();
    window.URL.revokeObjectURL(url);
  };

  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: loadCalendarData
    },
    {
      key: 'filters',
      text: 'Filters',
      iconProps: { iconName: 'Filter' },
      onClick: () => setState(prev => ({ ...prev, showFiltersPanel: true }))
    }
  ];

  if (allowExport) {
    commandBarItems.push({
      key: 'export',
      text: 'Export',
      iconProps: { iconName: 'Download' },
      onClick: exportToCSV
    });
  }

  const departmentOptions: IDropdownOption[] = [
    { key: 'All', text: 'All Departments' },
    { key: 'Engineering', text: 'Engineering' },
    { key: 'Marketing', text: 'Marketing' },
    { key: 'Sales', text: 'Sales' },
    { key: 'HR', text: 'Human Resources' }
  ];

  const statusOptions: IDropdownOption[] = [
    { key: 'All', text: 'All Statuses' },
    { key: 'Pending', text: 'Pending' },
    { key: 'Approved', text: 'Approved' },
    { key: 'Rejected', text: 'Rejected' }
  ];

  const listColumns: IColumn[] = [
    {
      key: 'employee',
      name: 'Employee',
      fieldName: 'requesterName',
      minWidth: 150,
      maxWidth: 200
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
      key: 'status',
      name: 'Status',
      fieldName: 'approvalStatus',
      minWidth: 80,
      maxWidth: 100
    }
  ];

  if (state.loading) {
    return (
      <div className={styles.teamCalendar}>
        <Stack horizontalAlign="center" verticalAlign="center" style={{ height: '200px' }}>
          <Spinner size={SpinnerSize.large} label="Loading calendar..." />
        </Stack>
      </div>
    );
  }

  return (
    <div className={`${styles.teamCalendar} ${isDarkTheme ? styles.dark : ''}`}>
      <Stack tokens={{ childrenGap: 20 }}>
        {title && (
          <Text variant="xxLarge" className={styles.title}>
            {title}
          </Text>
        )}
        
        {state.error && (
          <MessageBar messageBarType={MessageBarType.error}>
            {state.error}
          </MessageBar>
        )}

        <CommandBar items={commandBarItems} />

        <div className={styles.calendarContainer}>
          <FullCalendar
            ref={calendarRef}
            plugins={[dayGridPlugin, timeGridPlugin, interactionPlugin]}
            headerToolbar={{
              left: 'prev,next today',
              center: 'title',
              right: 'dayGridMonth,timeGridWeek,timeGridDay'
            }}
            initialView={defaultView}
            weekends={showWeekends}
            events={state.events}
            eventClick={handleEventClick}
            select={handleDateSelect}
            selectable={true}
            selectMirror={true}
            dayMaxEvents={true}
            height="auto"
            eventDisplay="block"
            displayEventTime={false}
          />
        </div>

        <Panel
          isOpen={state.showEventPanel}
          onDismiss={() => setState(prev => ({ ...prev, showEventPanel: false }))}
          type={PanelType.medium}
          headerText="Leave Request Details"
        >
          {state.selectedEvent && (
            <Stack tokens={{ childrenGap: 15 }}>
              <TextField label="Employee" value={state.selectedEvent.requesterName} readOnly />
              <TextField label="Leave Type" value={state.selectedEvent.leaveType} readOnly />
              <TextField label="Start Date" value={moment(state.selectedEvent.startDate).format('MMM DD, YYYY')} readOnly />
              <TextField label="End Date" value={moment(state.selectedEvent.endDate).format('MMM DD, YYYY')} readOnly />
              <TextField label="Total Days" value={(state.selectedEvent.totalDays || 0).toString()} readOnly />
              <TextField label="Status" value={state.selectedEvent.approvalStatus} readOnly />
              {state.selectedEvent.requestComments && (
                <TextField label="Comments" value={state.selectedEvent.requestComments} multiline rows={3} readOnly />
              )}
              <DefaultButton
                text="Close"
                onClick={() => setState(prev => ({ ...prev, showEventPanel: false }))}
              />
            </Stack>
          )}
        </Panel>

        <Panel
          isOpen={state.showFiltersPanel}
          onDismiss={() => setState(prev => ({ ...prev, showFiltersPanel: false }))}
          type={PanelType.medium}
          headerText="Filter Options"
        >
          <Stack tokens={{ childrenGap: 15 }}>
            <TextField
              label="Search"
              placeholder="Search by employee name or leave type"
              value={state.searchText}
              onChange={(_, newValue) => setState(prev => ({ ...prev, searchText: newValue || '' }))}
            />
            <Dropdown
              label="Department"
              options={departmentOptions}
              selectedKey={state.selectedDepartment}
              onChange={(_, option) => setState(prev => ({ ...prev, selectedDepartment: option?.key as string }))}
            />
            <Dropdown
              label="Status"
              options={statusOptions}
              selectedKey={state.selectedStatus}
              onChange={(_, option) => setState(prev => ({ ...prev, selectedStatus: option?.key as string }))}
            />
            <Text variant="medium">Filtered Results: {state.filteredRequests.length}</Text>
            <DetailsList
              items={state.filteredRequests.slice(0, 10)}
              columns={listColumns}
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              compact={true}
            />
            <DefaultButton
              text="Close"
              onClick={() => setState(prev => ({ ...prev, showFiltersPanel: false }))}
            />
          </Stack>
        </Panel>
      </Stack>
    </div>
  );
};

export default TeamCalendar;
