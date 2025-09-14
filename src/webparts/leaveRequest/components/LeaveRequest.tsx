import * as React from 'react';
//import styles from './LeaveRequest.module.scss';
import type { ILeaveRequestProps } from './ILeaveRequestProps';
import { getSP } from '../../../pnpjsConfig';
import { SPFI } from '@pnp/sp';
import { useEffect, useState } from 'react';
//import { escape } from '@microsoft/sp-lodash-subset';
import { Dropdown, DatePicker, PrimaryButton, TextField, DetailsList, DetailsListLayoutMode, SelectionMode, IDropdownOption } from '@fluentui/react';
import { AnimatedBackground } from 'animated-backgrounds';

const LeaveRequest = (props: ILeaveRequestProps) => {

  const sp: SPFI = getSP(props.context);
  const listname: string = "LeaveRequests"; //change listname here if needed
  // State variables for form fields and leave requests
  
  const [employeeName, setEmployeeName] = useState('');
  const [employeeEmail, setEmployeeEmail] = useState('');
  const [leaveType, setLeaveType] = useState('');
  const [startDate, setStartDate] = useState<Date | null>(null);
  const [endDate, setEndDate] = useState<Date | null>(null);
  const [approverEmail, setApproverEmail] = useState('');
  const [leaveRequests, setLeaveRequests] = useState<any[]>([]);
  const [showForm, setShowForm] = useState(false);
  const [selectedRequest, setSelectedRequest] = useState<any | null>(null); // Track selected item
  const loggedinuser = props.context.pageContext.user.email;
//useEffect to fetch leave requests on component mount
  useEffect(() => {
      fetchLeaveRequests();
  }, []);

  // Function to handle form submission
  const handleSubmit = async () => {
      try {
        if(employeeEmail.trim() === '' || leaveType.trim() === '' || !startDate || !endDate || approverEmail.trim() === ''){
          alert("Please fill all required fields");
          return;
        }
        if (startDate > endDate) {
          alert("Start Date cannot be later than End Date");
          return;
        }
          await sp.web.lists.getByTitle(listname).items.add({
              Title: employeeName,
              EmployeeEmail: employeeEmail,
              LeaveType: leaveType,
              StartDate: startDate,
              EndDate: endDate,
              Status: "Pending",
              ApproverEmail: approverEmail,
          });
          alert("Leave request submitted successfully");
          fetchLeaveRequests();
          setShowForm(false); // Return to list after submission
      } catch (error) {
          console.error("Error submitting request", error);
          alert("Error submitting request");
      }
  };

  // Function to fetch leave requests from SharePoint
  const fetchLeaveRequests = async () => {
      try {
          const items = await sp.web.lists.getByTitle(listname).items
              .filter(`EmployeeEmail eq '${loggedinuser}'`)
              .select("Title", "EmployeeEmail", "LeaveType", "StartDate", "EndDate", "Status", "ApproverEmail")();
          setLeaveRequests(items);
      } catch (error) {
          console.error("Error fetching leave requests", error);
      }
  };

  // Define columns for DetailsList
  const columns = [
      { key: 'LeaveType', name: 'Leave Type', fieldName: 'LeaveType', minWidth: 100, maxWidth: 150, isResizable: true },
      { key: 'StartDate', name: 'Start Date', fieldName: 'StartDate', minWidth: 100, maxWidth: 150, isResizable: true },
      { key: 'EndDate', name: 'End Date', fieldName: 'EndDate', minWidth: 100, maxWidth: 150, isResizable: true },
      { key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 100, maxWidth: 150, isResizable: true },
      { key: 'ApproverEmail', name: 'Approver Email', fieldName: 'ApproverEmail', minWidth: 100, maxWidth: 150, isResizable: true },
  ];

  // Handle row click to view details
  const handleRowClick = (item: any) => {
      setSelectedRequest(item);
      setEmployeeName(item.Title);
      setEmployeeEmail(item.EmployeeEmail);
      setLeaveType(item.LeaveType);
      setStartDate(new Date(item.StartDate));
      setEndDate(new Date(item.EndDate));
      setApproverEmail(item.ApproverEmail);
      setShowForm(true);
  };

  // Handle new request button click
  const handleNewRequest = () => {
      setSelectedRequest(null);
      setEmployeeName('');
      setEmployeeEmail('');
      setLeaveType('');
      setStartDate(null);
      setEndDate(null);
      setApproverEmail('');
      setShowForm(true);
  };

  // Render component
  return (
      <div>       
        <AnimatedBackground animationName="matrixRain" />
          <h2>Leave Requests</h2>
          <div style={{ marginBottom: '10px' }}>
              <PrimaryButton text="My Leave Requests" onClick={() => setShowForm(false)} />
              <PrimaryButton text="New Leave Request" onClick={handleNewRequest} style={{ marginLeft: '10px' }} />
          </div>

          {!showForm ? (
              <DetailsList
                  items={leaveRequests}
                  columns={columns}
                  setKey="set"
                  layoutMode={DetailsListLayoutMode.fixedColumns}
                  selectionMode={SelectionMode.single}
                  selectionPreservedOnEmptyClick={true}
                  onActiveItemChanged={handleRowClick} // Row selection handler
              />
          ) : (
              <div>
                  <h2>{selectedRequest ? "Leave Request Details" : "New Leave Request"}</h2>
                  <TextField label="Employee Name" value={employeeName} disabled={!!selectedRequest} onChange={(e, newValue) => setEmployeeName(newValue || '')} />
                  <TextField label="Employee Email" value={employeeEmail} disabled={!!selectedRequest} onChange={(e, newValue) => setEmployeeEmail(newValue || '')} required />
                  <Dropdown
                      label="Leave Type"
                      options={[{ key: 'Sick Leave', text: 'Sick Leave' }, { key: 'Vacation', text: 'Vacation' }, { key: 'Casual Leave', text: 'Casual Leave' }]}
                      selectedKey={leaveType}
                      disabled={!!selectedRequest}
                      onChange={(_, option?: IDropdownOption) => setLeaveType((option?.key as string) || '')}
                  />
                  <DatePicker label="Start Date" placeholder="Select a start date..." value={startDate || undefined} onSelectDate={(date) => setStartDate(date || null)} disabled={!!selectedRequest} isRequired={true}/>
                  <DatePicker label="End Date" placeholder="Select a end date..." value={endDate || undefined} onSelectDate={(date) => setEndDate(date || null)} disabled={!!selectedRequest} isRequired={true} />
                  <TextField label="Approver Email" value={approverEmail} disabled={!!selectedRequest} onChange={(e, newValue) => setApproverEmail(newValue || '')} required/>
                  {!selectedRequest && <PrimaryButton text="Submit Request" onClick={handleSubmit} />}
              </div>
          )}
      </div>
  );
};

export default LeaveRequest;