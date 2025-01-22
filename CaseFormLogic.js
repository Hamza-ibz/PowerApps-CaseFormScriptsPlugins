
"use strict";

/**
 * Main function executed when the form loads.
 * @param {Object} executionContext - The execution context of the form.
 */
function onFormLoad(executionContext) {
    try {
        const formContext = getFormContext(executionContext);
        displayWelcomeNotification(formContext);
        handleCustomerField(formContext);
    } catch (error) {
        handleFormLoadError(error, executionContext);
    }
}

function getFormContext(executionContext) {
    return executionContext.getFormContext();
}

function displayWelcomeNotification(formContext) {
    formContext.ui.setFormNotification("Hello world v16", "INFO", "IDUnique220912");
}

function handleCustomerField(formContext) {
    const customerField = formContext.getAttribute("customerid");
    if (customerField && customerField.getValue()) {
        const customerData = customerField.getValue()[0];
        const customerId = customerData.id.replace(/[{}]/g, ""); // Remove curly braces
        const customerType = customerData.entityType;

        handleCustomerType(formContext, customerType, customerId);
    } else {
        handleNoCustomerSelected(formContext);
    }
}

function handleCustomerType(formContext, customerType, customerId) {
    if (customerType === "account") {
        handleAccountCustomer(formContext, customerId);
    } else if (customerType === "contact") {
        handleContactCustomer(formContext);
    }
}

function handleNoCustomerSelected(formContext) {
    setFieldVisibility(formContext); // Ensure fields are adjusted for no customer
    resetContactField(formContext); // Clear contact field if no customer
}

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
function handleAccountCustomer(formContext, accountId) {
    retrieveAccountDetails(accountId)
        .then((account) => {
            ensureContactFieldVisibility(formContext);
            setContactFieldRequirement(formContext, "required");

            if (account._primarycontactid_value) {
                const contactId = account._primarycontactid_value;
                retrieveContactDetails(formContext, contactId);
            } else {
                handleNoPrimaryContact(formContext);
            }
        })
        .catch((error) => {
            handleAccountError(formContext, error);
        });
}

function retrieveAccountDetails(accountId) {
    return Xrm.WebApi.retrieveRecord("account", accountId, "?$select=_primarycontactid_value");
}

function ensureContactFieldVisibility(formContext) {
    const contactControl = formContext.getControl("primarycontactid");
    if (contactControl) {
        contactControl.setVisible(true);
    }
}

function setContactFieldRequirement(formContext, requirementLevel) {
    const contactAttribute = formContext.getAttribute("primarycontactid");
    if (contactAttribute) {
        contactAttribute.setRequiredLevel(requirementLevel);
    }
}

function handleNoPrimaryContact(formContext) {
    console.warn("No primary contact associated with this account.");
    formContext.ui.setFormNotification(
        "No primary contact is associated with this account.",
        "WARNING",
        "NoPrimaryContact"
    );
    setFieldVisibility(formContext); // Ensure visibility adjustments as necessary
}

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
function retrieveContactDetails(formContext, contactId) {
    Xrm.WebApi.retrieveRecord("contact", contactId, "?$select=emailaddress1,mobilephone,fullname")
        .then((contact) => {
            console.log("Contact Record Retrieved: ", contact);
            setFieldVisibility(formContext); // Updates visibility in quick view
            formContext.getAttribute("primarycontactid").setValue([
                { id: contactId, name: contact.fullname, entityType: "contact" }
            ]);
        })
        .catch((error) => {
            console.error("Error retrieving contact details: ", error);
            formContext.ui.setFormNotification(
                "Unable to retrieve contact details. Please try again later.",
                "ERROR",
                "ContactError"
            );
            setFieldVisibility(formContext);
        });
}

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

function getQuickViewControl(formContext, controlName) {
    const quickViewControl = formContext.ui.quickForms.get(controlName);
    if (!quickViewControl) {
        console.warn(`${controlName} not found!`);
        return null;
    }
    return quickViewControl;
}

function waitForQuickViewLoad(quickViewControl, callback) {
    const waitForLoad = setInterval(() => {
        if (quickViewControl.isLoaded()) {
            clearInterval(waitForLoad);
            callback();
        }
    }, 500);
}

function getFieldValue(quickViewControl, fieldName) {
    try {
        return quickViewControl.getControl(fieldName).getAttribute().getValue()?.trim() || null;
    } catch (error) {
        console.warn(`Field ${fieldName} not found or inaccessible.`);
        return null;
    }
}

function updateFieldVisibility(quickViewControl, fieldName, value) {
    try {
        quickViewControl.getControl(fieldName).setVisible(!!value);
    } catch (error) {
        console.warn(`Unable to update visibility for field ${fieldName}:`, error);
    }
}

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
 * Sets the contact field requirement level.
 * @param {Object} formContext - The form context.
 * @param {string} level - "required" or "none".
 */
function setContactFieldRequirement(formContext, level) {
    const contactField = formContext.getAttribute("primarycontactid");
    if (contactField) {
        contactField.setRequiredLevel(level);
    }
}

/**
 * Clears the contact field value.
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











