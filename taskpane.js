// Wait for Office to be ready
Office.onReady(() => {
    if (Office.context.mailbox) {
        console.log("Add-in is ready!");
    }
});

// Function to display email attachments
async function showAttachments() {
    const item = Office.context.mailbox.item;

    // Check if there are any attachments
    if (item.attachments && item.attachments.length > 0) {
        const attachmentList = document.getElementById("attachmentList");
        attachmentList.innerHTML = ""; // Clear previous list

        // Loop through attachments and display their names
        item.attachments.forEach((attachment) => {
            const listItem = document.createElement("li");
            listItem.textContent = attachment.name;
            attachmentList.appendChild(listItem);
        });
    } else {
        // No attachments found
        document.getElementById("attachmentList").textContent = "No attachments found.";
    }
}

// Bind the button to the function
document.getElementById("showAttachmentsButton").onclick = showAttachments;
