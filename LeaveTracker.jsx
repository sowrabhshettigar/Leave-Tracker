import React, { useState, useEffect } from "react";
import {
  TextField,
  Dropdown,
  DatePicker,
  PrimaryButton,
  Spinner,
  SpinnerSize,
} from "@fluentui/react";
import { PeoplePicker } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

const LeaveTracker = (props) => {
  const { context, webAbsoluteUrl } = props;

  const [formState, setFormState] = useState({
    employeeName: "",
    leaveReason: "",
    leaveType: "",
    startDate: null,
    endDate: null,
    approvers: [],
  });

  const [leaveTypes, setLeaveTypes] = useState([]);
  const [loadingLeaveTypes, setLoadingLeaveTypes] = useState(true);
  const [loadingItem, setLoadingItem] = useState(false);
  const [isUpdateMode, setIsUpdateMode] = useState(false);
  const [itemIdInput, setItemIdInput] = useState("");
  const [itemId, setItemId] = useState(null);

  // Fetch Leave Types
  useEffect(() => {
    const fetchLeaveTypes = async () => {
      try {
        setLoadingLeaveTypes(true);
        const sp = spfi(webAbsoluteUrl).using(SPFx(context));
        const items = await sp.web.lists.getByTitle("Leave Types").items();
        setLeaveTypes(
          items.map((type) => ({
            key: type.ID,
            text: type.Title,
          }))
        );
      } catch (error) {
        console.error("Failed to load leave types:", error);
        alert(`Failed to load leave types. Details: ${error.message || error}`);
      } finally {
        setLoadingLeaveTypes(false);
      }
    };

    fetchLeaveTypes();
  }, [context, webAbsoluteUrl]);

  // Fetch item data when Item ID is provided
  const handleFetchForUpdate = async () => {
    try {
      if (!itemIdInput) {
        alert("Please enter a valid Item ID.");
        return;
      }

      setLoadingItem(true);
      const sp = spfi(webAbsoluteUrl).using(SPFx(context));
      const item = await sp.web.lists
        .getByTitle("Leave Requests")
        .items.getById(itemIdInput)
        .select("*", "Approvers/Id", "Approvers/Title", "LeaveType/Id", "LeaveType/Title")
        .expand("Approvers", "LeaveType")();

      setFormState({
        employeeName: item.EmployeeName || "",
        leaveReason: item.ReasonforLeave || "",
        leaveType: item.LeaveType?.Id || "",
        startDate: item.StartDate ? new Date(item.StartDate) : null,
        endDate: item.EndDate ? new Date(item.EndDate) : null,
        approvers: item.Approvers
          ? [
              {
                id: item.Approvers.Id,
                text: item.Approvers.Title,
              },
            ]
          : [],
      });
      setItemId(itemIdInput); // Set the current item ID for updating
      alert("Form populated successfully! You can now update the request.");
    } catch (error) {
      // Check for 404 (Item not found) error
      if (error?.response?.status === 404) {
        alert("Form with this ID does not exist.");
      } else {
        console.error("Failed to load item data:", error);
        alert(`Failed to load item data. Details: ${error.message || error}`);
      }
    } finally {
      setLoadingItem(false);
    }
  };

  // Handle form submit (Create or Update)
  const handleSubmit = async () => {
    try {
      if (
        !formState.employeeName ||
        !formState.leaveReason ||
        !formState.leaveType ||
        !formState.startDate ||
        !formState.endDate
      ) {
        alert("Please fill in all required fields.");
        return;
      }

      const formattedStartDate = formState.startDate
        ? formState.startDate.toISOString().split("T")[0]
        : null;
      const formattedEndDate = formState.endDate
        ? formState.endDate.toISOString().split("T")[0]
        : null;

      const sp = spfi(webAbsoluteUrl).using(SPFx(context));
      const itemToSubmit = {
        Title: formState.employeeName || "",
        EmployeeName: formState.employeeName || "",
        ReasonforLeave: formState.leaveReason || "",
        LeaveTypeId: formState.leaveType,
        StartDate: formattedStartDate,
        EndDate: formattedEndDate,
      };

      if (formState.approvers && formState.approvers.length > 0) {
        const approverData = formState.approvers[0];
        if (approverData) {
          itemToSubmit.ApproversId = approverData.id || approverData.key;
        }
      }

      if (itemId) {
        // Update existing item
        await sp.web.lists
          .getByTitle("Leave Requests")
          .items.getById(itemId)
          .update(itemToSubmit);
        alert("Leave request updated successfully!");
      } else {
        // Create new item
        const response = await sp.web.lists
          .getByTitle("Leave Requests")
          .items.add(itemToSubmit);
          const newItemId = response.Id || response.data?.Id;
          // Get the newly created Item ID
        alert(`Leave request submitted successfully! Your Item ID is: ${newItemId}`);
      }

      // Reset form after successful submission (if new item)
      if (!itemId) {
        setFormState({
          employeeName: "",
          leaveReason: "",
          leaveType: "",
          startDate: null,
          endDate: null,
          approvers: [],
        });
      }
    } catch (error) {
      console.error("Error submitting leave request:", error);
      alert(`Error: ${error.message || error}`);
    }
  };

  // Show loading spinner if loading item data
  if (loadingItem) {
    return <Spinner label="Loading leave request data..." size={SpinnerSize.large} />;
  }

  return (
    <div>
      <h1>{itemId ? "Edit Leave Request" : "New Leave Request"}</h1>

      {isUpdateMode && (
        <div>
          <TextField
            label="Enter Item ID to Update"
            value={itemIdInput}
            onChange={(e, newValue) => setItemIdInput(newValue)}
          />
          <PrimaryButton
            text="Fetch Details"
            onClick={handleFetchForUpdate}
            style={{ marginBottom: 20 }}
          />
        </div>
      )}

      <TextField
        label="Employee Name"
        value={formState.employeeName}
        onChange={(e, newValue) => setFormState({ ...formState, employeeName: newValue })}
        required
      />

      <TextField
        label="Reason for Leave"
        value={formState.leaveReason}
        onChange={(e, newValue) => setFormState({ ...formState, leaveReason: newValue })}
        required
      />

      {loadingLeaveTypes ? (
        <Spinner label="Loading leave types..." size={SpinnerSize.medium} />
      ) : (
        <Dropdown
          label="Leave Type"
          options={leaveTypes}
          selectedKey={formState.leaveType}
          onChange={(e, option) => setFormState({ ...formState, leaveType: option?.key })}
          required
        />
      )}

      <DatePicker
        label="Start Date"
        onSelectDate={(date) => setFormState({ ...formState, startDate: date })}
        value={formState.startDate}
        required
      />

      <DatePicker
        label="End Date"
        onSelectDate={(date) => setFormState({ ...formState, endDate: date })}
        value={formState.endDate}
        required
      />

      <PeoplePicker
        context={context}
        titleText="Approvers"
        personSelectionLimit={1}
        showtooltip={true}
        required={true}
        onChange={(items) => setFormState({ ...formState, approvers: items })}
        ensureUser={true}
        webAbsoluteUrl={webAbsoluteUrl}
        defaultSelectedUsers={formState.approvers.map((a) => a.text)}
      />

      <PrimaryButton
        text={itemId ? "Update" : "Submit"}
        onClick={handleSubmit}
        style={{ marginTop: 20 }}
      />
      <PrimaryButton
        text={isUpdateMode ? "Switch to Create Mode" : "Switch to Update Mode"}
        onClick={() => {
          setIsUpdateMode(!isUpdateMode);
          setItemId(null);
          setFormState({
            employeeName: "",
            leaveReason: "",
            leaveType: "",
            startDate: null,
            endDate: null,
            approvers: [],
          });
        }}
        style={{ marginLeft: 10, marginTop: 20 }}
      />
    </div>
  );
};

export default LeaveTracker;
