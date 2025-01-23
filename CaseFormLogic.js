"use strict";

/**
 * Main function executed when the form loads.
 * @param {Object} executionContext - The execution context of the form.
 */
async function onFormLoad(executionContext) {
    try {
        const formContext = executionContext.getFormContext();
        await handleCustomerField(formContext);
    } catch (error) {
        handleFormLoadError(error, executionContext);
    }
}

/**
 * Handles logic related to the customer field (account or contact).
 * @param {Object} formContext - The form context.
 */
async function handleCustomerField(formContext) {
    const customerField = formContext.getAttribute("customerid");
    if (customerField && customerField.getValue()) {
        const customerData = customerField.getValue()[0];
        const customerId = customerData.id.replace(/[{}]/g, ""); // Remove curly braces
        const customerType = customerData.entityType;

        await handleCustomerType(formContext, customerType, customerId);
    } else {
        handleNoCustomerSelected(formContext);
    }
}

/**
 * Handles logic for customer type (Account or Contact).
 * @param {Object} formContext - The form context.
 * @param {string} customerType - The type of customer (account or contact).
 * @param {string} customerId - The ID of the customer.
 */
async function handleCustomerType(formContext, customerType, customerId) {
    if (customerType === "account") {
        await handleAccountCustomer(formContext, customerId);
    } else if (customerType === "contact") {
        handleContactCustomer(formContext);
    }
}

/**
 * Handles cases where no customer is selected.
 * @param {Object} formContext - The form context.
 */
function handleNoCustomerSelected(formContext) {
    setFieldVisibility(formContext); // Ensure fields are adjusted for no customer
    resetContactField(formContext); // Clear contact field if no customer
}

/**
 * Handles errors during form load.
 * @param {Error} error - The error object.
 * @param {Object} executionContext - The execution context of the form.
 */
function handleFormLoadError(error, executionContext) {
    console.error("Error in onFormLoad: ", error);
    notifyUserOfError(error, executionContext);
}

/**
 * Handles logic when the customer is an Account.
 * Ensures the contact field is visible and mandatory.
 * @param {Object} formContext - The form context.
 * @param {string} accountId - The ID of the Account.
 */
async function handleAccountCustomer(formContext, accountId) {
    try {
        const account = await Xrm.WebApi.retrieveRecord("account", accountId, "?$select=_primarycontactid_value");;
        ensureContactFieldVisibility(formContext);
        setContactFieldRequirement(formContext, "required");

        if (account._primarycontactid_value) {
            const contactId = account._primarycontactid_value;
            await retrieveContactDetails(formContext, contactId);
        } else {
            handleNoPrimaryContact(formContext);
        }
    } catch (error) {
        handleAccountError(formContext, error);
    }
}

/**
 * Ensures the contact field is visible on the form.
 * @param {Object} formContext - The form context.
 */
function ensureContactFieldVisibility(formContext) {
    const contactControl = formContext.getControl("primarycontactid");
    if (contactControl) {
        contactControl.setVisible(true);
    }
}

/**
 * Sets the contact field requirement level.
 * @param {Object} formContext - The form context.
 * @param {string} requirementLevel - The requirement level ("required" or "none").
 */
function setContactFieldRequirement(formContext, requirementLevel) {
    const contactAttribute = formContext.getAttribute("primarycontactid");
    if (contactAttribute) {
        contactAttribute.setRequiredLevel(requirementLevel);
    }
}

/**
 * Handles cases where no primary contact is associated with the account.
 * @param {Object} formContext - The form context.
 */
function handleNoPrimaryContact(formContext) {
    console.warn("No primary contact associated with this account.");
    formContext.ui.setFormNotification(
        "No primary contact is associated with this account.",
        "WARNING",
        "NoPrimaryContact"
    );
    setFieldVisibility(formContext); // Ensure visibility adjustments as necessary
}

/**
 * Handles errors during account processing.
 * @param {Object} formContext - The form context.
 * @param {Error} error - The error object.
 */
function handleAccountError(formContext, error) {
    console.error("Error retrieving account details: ", error);
    formContext.ui.setFormNotification(
        "Unable to retrieve account details. Please try again later.",
        "ERROR",
        "AccountError"
    );
    setFieldVisibility(formContext); // Ensure visibility adjustments as necessary
}

/**
 * Handles logic when the customer is a Contact.
 * Hides the contact field since the customer is a contact.
 * @param {Object} formContext - The form context.
 */
function handleContactCustomer(formContext) {
    const contactControl = formContext.getControl("primarycontactid");
    if (contactControl) {
        contactControl.setVisible(false); // Hide contact field
    }
    setContactFieldRequirement(formContext, "none"); // Make contact field optional
}

/**
 * Retrieves the details of a Contact and updates the form fields accordingly.
 * @param {Object} formContext - The form context.
 * @param {string} contactId - The ID of the Contact.
 */
async function retrieveContactDetails(formContext, contactId) {
    try {
        const contact = await Xrm.WebApi.retrieveRecord(
            "contact",
            contactId,
            "?$select=emailaddress1,mobilephone,fullname"
        );

        console.log("Contact Record Retrieved: ", contact);
        setFieldVisibility(formContext); // Updates visibility in quick view
        formContext.getAttribute("primarycontactid").setValue([
            { id: contactId, name: contact.fullname, entityType: "contact" }
        ]);
    } catch (error) {
        console.error("Error retrieving contact details: ", error);
        formContext.ui.setFormNotification(
            "Unable to retrieve contact details. Please try again later.",
            "ERROR",
            "ContactError"
        );
        setFieldVisibility(formContext);
    }
}

/**
 * Adjusts the visibility of the quick view control.
 * @param {Object} formContext - The form context.
 */
function setFieldVisibility(formContext) {
    try {
        const quickViewControl = getQuickViewControl(formContext, "ContactDetailsQuickView");
        if (!quickViewControl) return;

        waitForQuickViewLoad(quickViewControl, () => {
            const email = getFieldValue(quickViewControl, "emailaddress1");
            const phone = getFieldValue(quickViewControl, "mobilephone");

            updateFieldVisibility(quickViewControl, "emailaddress1", email);
            updateFieldVisibility(quickViewControl, "mobilephone", phone);

            const shouldHideQuickView = !email && !phone;
            quickViewControl.setVisible(!shouldHideQuickView);

            updateFormNotification(formContext, shouldHideQuickView);
        });
    } catch (error) {
        console.error("Error in setting Quick View Form visibility:", error);
    }
}

/**
 * Retrieves a Quick View control by name.
 * @param {Object} formContext - The form context.
 * @param {string} controlName - The name of the Quick View control.
 * @returns {Object|null} - The Quick View control or null if not found.
 */

function getQuickViewControl(formContext, controlName) {
    const quickViewControl = formContext.ui.quickForms.get(controlName);
    if (!quickViewControl) {
        console.warn(`${controlName} not found!`);
        return null;
    }
    return quickViewControl;
}

/**
 * Waits for a Quick View control to finish loading before executing a callback function.
 * @param {Object} quickViewControl - The Quick View control to monitor.
 * @param {Function} callback - The function to execute once the Quick View is loaded.
 */
function waitForQuickViewLoad(quickViewControl, callback) {
    const waitForLoad = setInterval(() => {
        if (quickViewControl.isLoaded()) {
            clearInterval(waitForLoad);
            callback();
        }
    }, 500);
}

/**
 * Retrieves the value of a specific field from a Quick View control.
 * @param {Object} quickViewControl - The Quick View control containing the field.
 * @param {string} fieldName - The logical name of the field.
 * @returns {string|null} - The trimmed field value if available, otherwise null.
 */
function getFieldValue(quickViewControl, fieldName) {
    try {
        return quickViewControl.getControl(fieldName).getAttribute().getValue()?.trim() || null;
    } catch (error) {
        console.warn(`Field ${fieldName} not found or inaccessible.`);
        return null;
    }
}

/**
 * Updates the visibility of a field within a Quick View control.
 * @param {Object} quickViewControl - The Quick View control containing the field.
 * @param {string} fieldName - The logical name of the field.
 * @param {boolean} value - The visibility status (true for visible, false for hidden).
 */
function updateFieldVisibility(quickViewControl, fieldName, value) {
    try {
        quickViewControl.getControl(fieldName).setVisible(!!value);
    } catch (error) {
        console.warn(`Unable to update visibility for field ${fieldName}:`, error);
    }
}

/**
 * Updates the form notification based on whether the Quick View control should be hidden.
 * @param {Object} formContext - The form context.
 * @param {boolean} shouldHideQuickView - Whether the Quick View should be hidden.
 */
function updateFormNotification(formContext, shouldHideQuickView) {
    const notificationId = "NoContactDetails";
    if (shouldHideQuickView) {
        formContext.ui.setFormNotification(
            "No contact details are available to display.",
            "WARNING",
            notificationId
        );
    } else {
        formContext.ui.clearFormNotification(notificationId);
    }
}

/**
 * Resets the contact field value to null.
 * @param {Object} formContext - The form context.
 */
function resetContactField(formContext) {
    const contactField = formContext.getAttribute("primarycontactid");
    if (contactField) {
        contactField.setValue(null);
    }
}

/**
 * Notifies the user of an error and logs the error details.
 * @param {Object} error - The error object.
 * @param {Object} executionContext - The execution context of the form.
 */
function notifyUserOfError(error, executionContext) {
    const formContext = executionContext.getFormContext();
    const message = "An error occurred. Please reload the form or contact an administrator.";
    formContext.ui.setFormNotification(message, "ERROR", "ErrorNotification");
    console.error("Error: ", error);
}


