/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log("Office is ready, initializing the app...");
    initializeApp();
  } else {
    console.error("This add-in is not running in Word.");
  }
});

async function initializeApp() {
  try {
    console.log("Initializing the app...");

    // Ensure elements exist before modifying them
    const sideloadMsg = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");

    if (sideloadMsg) sideloadMsg.style.display = "none";
    if (appBody) appBody.style.display = "flex";

    // Populate dropdown with dummy tags (clear existing first)
    await populateDummyTags();

    // Fix: Ensure we remove any previous event listeners to prevent multiple inserts
    const button = document.getElementById("insertTagButton");
    button.replaceWith(button.cloneNode(true)); // Removes existing event listeners
    document.getElementById("insertTagButton").addEventListener("click", insertSelectedTag);

    console.log("App initialized successfully.");
  } catch (error) {
    console.error("Error initializing app:", error);
  }
}

// Populate the dropdown with dummy tags
function populateDummyTags() {
  const dummyTags = [
    "{{FirstName}}",
    "{{LastName}}",
    "{{Date}}",
    "{{Email}}",
    "{{PhoneNumber}}",
    "{{Address}}",
    "{{Position}}"
  ];

  const dropdown = document.getElementById("tagsDropdown");
  if (!dropdown) {
    console.error("Dropdown element not found.");
    return;
  }

  // Fix: Clear previous dropdown values before inserting new ones
  dropdown.innerHTML = '<option value="" disabled selected>Select a tag</option>';

  dummyTags.forEach((tag) => {
    const option = document.createElement("option");
    option.value = tag;
    option.textContent = tag;
    dropdown.appendChild(option);
  });

  console.log("Dummy tags added to dropdown:", dummyTags);
}

// Insert the selected tag into the Word document
function insertSelectedTag() {
  const dropdown = document.getElementById("tagsDropdown");
  const selectedTag = dropdown.value;

  if (!selectedTag) {
    console.warn("No tag selected.");
    return;
  }

  console.log("Inserting tag into document:", selectedTag);

  Word.run((context) => {
    const selection = context.document.getSelection();
    
    // Fix: Ensure that only one tag is inserted by logging before execution
    console.log("Final tag to be inserted:", selectedTag);

    selection.insertText(selectedTag, Word.InsertLocation.replace);

    return context.sync().then(() => {
      console.log("Tag inserted successfully:", selectedTag);
    });
  }).catch((error) => {
    console.error("Error inserting tag:", error);
  });
}
