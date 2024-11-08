// This function creates a custom menu in the Google Sheets UI
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Square API')
    .addItem('Set API Key', 'setApiKey') // Menu item to set the API key
    .addItem('Set Email Address', 'setEmailAddress') // Menu item to set the email address
    .addItem('Start Processing', 'startProcessing')
    .addItem('Set Daily Timer', 'createDailyTrigger')
    .addToUi();
}

// Function to prompt the user to enter their Square API key
function setApiKey() {
  var ui = SpreadsheetApp.getUi();

  // Prompt the user for the API key
  var response = ui.prompt('Set Square API Key', 'Please enter your Square API access token:', ui.ButtonSet.OK_CANCEL);

  // Process the user's response
  if (response.getSelectedButton() == ui.Button.OK) {
    var apiKey = response.getResponseText().trim();

    if (apiKey) {
      // Store the API key in document properties
      var documentProperties = PropertiesService.getDocumentProperties();
      documentProperties.setProperty('SQUARE_ACCESS_TOKEN', apiKey);

      // Inform the user that the key has been saved
      ui.alert('Success', 'Your Square API access token has been saved securely.', ui.ButtonSet.OK);
    } else {
      ui.alert('Error', 'No API key entered. Please try again.', ui.ButtonSet.OK);
    }
  } else {
    ui.alert('Operation cancelled.');
  }
}

// Function to prompt the user to enter their email address
function setEmailAddress() {
  var ui = SpreadsheetApp.getUi();

  // Prompt the user for the email address
  var response = ui.prompt('Set Notification Email', 'Please enter your email address:', ui.ButtonSet.OK_CANCEL);

  // Process the user's response
  if (response.getSelectedButton() == ui.Button.OK) {
    var emailAddress = response.getResponseText().trim();

    if (emailAddress) {
      // Store the email address in document properties
      var documentProperties = PropertiesService.getDocumentProperties();
      documentProperties.setProperty('NOTIFICATION_EMAIL', emailAddress);

      // Inform the user that the email address has been saved
      ui.alert('Success', 'Your email address has been saved.', ui.ButtonSet.OK);
    } else {
      ui.alert('Error', 'No email address entered. Please try again.', ui.ButtonSet.OK);
    }
  } else {
    ui.alert('Operation cancelled.');
  }
}

// Main function to start processing
function startProcessing() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'API-Export';

  try {
    // Clear existing sheet if it exists, or create a new one if it doesn't
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.clear(); // Clear the content but keep the sheet
    } else {
      sheet = ss.insertSheet();
      sheet.setName(sheetName);
    }

    // Similar steps for the Progress sheet
    var progressSheetName = 'Processing-Progress';
    var progressSheet = ss.getSheetByName(progressSheetName);
    if (progressSheet) {
      progressSheet.clear(); // Clear the content but keep the sheet
    } else {
      progressSheet = ss.insertSheet();
      progressSheet.setName(progressSheetName);
    }

    // Initialize progress indicators
    progressSheet.getRange('A1').setValue('Total Variations:');
    progressSheet.getRange('A2').setValue('Variations Processed:');
    progressSheet.getRange('A3').setValue('Progress (%):');
    progressSheet.getRange('A5').setValue('Type "STOP" in cell B5 to halt processing.');

    // Reset stop flag
    progressSheet.getRange('B5').setValue('');

    // Fetch all location IDs automatically
    var locationIds = fetchLocationIds();
    if (!locationIds.length) {
      Logger.log("No locations found for the merchant.");
      displayAlert("No locations found for this merchant.");
      return;
    }

    // Fetch all catalog items and variations
    var catalogData = fetchAllCatalogItems();
    if (catalogData.items.length === 0) {
      displayAlert("No items found in the catalog.");
      return;
    }

    // Fetch all categories
    var categoryMap = fetchAllCategories();

    // Update total variations in progress sheet
    progressSheet.getRange('B1').setValue(catalogData.variationCount);

    // Create a variation map
    var variationMap = buildVariationMap(catalogData.items, categoryMap);

    // Fetch inventory counts for all variations
    var inventoryMap = fetchInventoryCountsForAllVariations(variationMap, locationIds, progressSheet);

    // Process variations and write data to the sheet
    processAndWriteData(sheet, variationMap, locationIds, inventoryMap, progressSheet);

    // Send success email
    var documentProperties = PropertiesService.getDocumentProperties();
    var emailAddress = documentProperties.getProperty('NOTIFICATION_EMAIL');

    if (emailAddress) {
      MailApp.sendEmail({
        to: emailAddress,
        subject: "Square Data Refresh Successful",
        body: "The daily refresh of Square data was completed successfully."
      });
    } else {
      Logger.log('Notification email address is not set. Please use "Set Email Address" in the "Square API" menu.');
    }

  } catch (error) {
    // Send failure email
    var documentProperties = PropertiesService.getDocumentProperties();
    var emailAddress = documentProperties.getProperty('NOTIFICATION_EMAIL');

    if (emailAddress) {
      MailApp.sendEmail({
        to: emailAddress,
        subject: "Square Data Refresh Failed",
        body: "The daily Square data refresh failed with the following error: " + error.message
      });
    } else {
      Logger.log('Notification email address is not set. Please use "Set Email Address" in the "Square API" menu.');
    }

    // Log the error in the Spreadsheet
    Logger.log("Error: " + error.message);
    // Optionally, show an alert if running manually
    displayAlert("An error occurred: " + error.message);
  }
}

// Function to create a time-driven trigger to refresh data daily
function createDailyTrigger() {
  // First, delete any existing triggers to avoid duplicates
  deleteExistingTriggers();

  // Set a time-driven trigger to run the startProcessing function every day at 8 AM
  ScriptApp.newTrigger('startProcessing')
    .timeBased()
    .atHour(8)  // Set the time here (8 AM in this case)
    .everyDays(1)  // Run every day
    .create();

  // Show confirmation to the user
  SpreadsheetApp.getUi().alert("Daily timer has been set for 8 AM.");
}

// Function to delete any existing time-based triggers for 'startProcessing' to avoid duplicates
function deleteExistingTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == 'startProcessing') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// Function to fetch all catalog items and variations
function fetchAllCatalogItems() {
  var allItems = [];
  var variationCount = 0;
  var cursor = null;
  var listCatalogUrl = 'https://connect.squareup.com/v2/catalog/list';

  do {
    var response = fetchCatalogPage(listCatalogUrl, cursor);

    if (response.getResponseCode() === 200) {
      var jsonData = JSON.parse(response.getContentText());

      if (Array.isArray(jsonData.objects) && jsonData.objects.length > 0) {
        jsonData.objects.forEach(function(item) {
          if (item.type === 'ITEM') {
            allItems.push(item);

            if (Array.isArray(item.item_data.variations)) {
              variationCount += item.item_data.variations.length;
            }
          }
        });
      }

      cursor = jsonData.cursor || null;
    } else {
      Logger.log("Error details from List Catalog: " + response.getContentText());
      displayAlert("Error retrieving catalog. Check logs for details.");
      return { items: [], variationCount: 0 };
    }

  } while (cursor);

  Logger.log("Total Items Retrieved: " + allItems.length);
  Logger.log("Total Variations Retrieved: " + variationCount);

  return { items: allItems, variationCount: variationCount };
}

// Function to fetch a catalog page
function fetchCatalogPage(listCatalogUrl, cursor) {
  var headers = {
    "Square-Version": "2023-10-18",
    "Content-Type": "application/json"
  };

  var urlWithCursor = cursor ? listCatalogUrl + "?cursor=" + cursor + "&include_related_objects=true" : listCatalogUrl + "?include_related_objects=true";

  var listOptions = {
    "method": "GET",
    "headers": headers,
    "muteHttpExceptions": true
  };

  return makeApiRequest(urlWithCursor, listOptions);
}

// Function to fetch all categories and build a category map
function fetchAllCategories() {
  var allCategories = {};
  var cursor = null;
  var listCatalogUrl = 'https://connect.squareup.com/v2/catalog/list';

  do {
    var urlWithCursor = cursor ? listCatalogUrl + "?cursor=" + cursor + "&types=CATEGORY" : listCatalogUrl + "?types=CATEGORY";
    var headers = {
      "Square-Version": "2023-10-18",
      "Content-Type": "application/json"
    };
    var options = {
      "method": "GET",
      "headers": headers,
      "muteHttpExceptions": true
    };
    var response = makeApiRequest(urlWithCursor, options);

    if (response.getResponseCode() === 200) {
      var jsonData = JSON.parse(response.getContentText());

      if (Array.isArray(jsonData.objects) && jsonData.objects.length > 0) {
        jsonData.objects.forEach(function(obj) {
          if (obj.type === "CATEGORY") {
            allCategories[obj.id] = obj.category_data.name;
          }
        });
      }

      cursor = jsonData.cursor || null;
    } else {
      Logger.log("Error fetching categories: " + response.getContentText());
      displayAlert("Error retrieving categories. Check logs for details.");
      return {};
    }

  } while (cursor);

  Logger.log("Total Categories Retrieved: " + Object.keys(allCategories).length);

  return allCategories;
}

// Function to build a variation map
function buildVariationMap(items, categoryMap) {
  var variationMap = {};

  items.forEach(function(item) {
    if (item.item_data && Array.isArray(item.item_data.variations)) {
      var itemId = item.id || "";
      var itemName = item.item_data.name || "";
      var description = item.item_data.description || "";
      var itemUrl = item.item_data.ecom_uri || "";

      // Get Image URLs (primary and additional)
      var imageUrls = [];
      if (Array.isArray(item.item_data.ecom_image_uris) && item.item_data.ecom_image_uris.length > 0) {
        imageUrls = item.item_data.ecom_image_uris; // Use ecom_image_uris first
      } else if (Array.isArray(item.image_ids) && item.image_ids.length > 0) {
        imageUrls = getImageUrls(item.image_ids, item.related_objects || []);  // Fallback to image_ids
      }

      // Ensure the columns for image URLs
      var primaryImageUrl = imageUrls[0] || "";
      var secondaryImageUrl = imageUrls[1] || "";
      var tertiaryImageUrl = imageUrls[2] || "";

      // Additional fields to capture
      var isDeleted = item.is_deleted || false;
      var catalogV1Ids = Array.isArray(item.catalog_v1_ids) ? item.catalog_v1_ids.map(function(id) {
        return id.catalog_v1_id;
      }).join(", ") : "";
      var itemVisibility = item.item_data.visibility || "";
      var categoryId = item.item_data.category_id || "";
      var categoryName = categoryMap[categoryId] || "";
      var modifierListInfo = item.item_data.modifier_list_info ? JSON.stringify(item.item_data.modifier_list_info) : "";
      var productType = item.item_data.product_type || "";
      var skipModifierScreen = item.item_data.skip_modifier_screen || false;
      var taxIds = Array.isArray(item.item_data.tax_ids) ? item.item_data.tax_ids.join(", ") : "";
      var itemOptions = item.item_data.item_options ? JSON.stringify(item.item_data.item_options) : "";
      var availableOnline = item.item_data.available_online || false;
      var availableForPickup = item.item_data.available_for_pickup || false;

      // Item-level availability fields
      var itemPresentAtAllLocations = item.hasOwnProperty('present_at_all_locations') ? item.present_at_all_locations : false;
      var itemPresentAtLocationIds = Array.isArray(item.present_at_location_ids) ? item.present_at_location_ids : [];
      var itemAbsentAtLocationIds = Array.isArray(item.absent_at_location_ids) ? item.absent_at_location_ids : [];

      item.item_data.variations.forEach(function(variation) {
        var variationId = variation.id || "";
        var variationName = variation.item_variation_data.name || "";
        var price = variation.item_variation_data.price_money
          ? variation.item_variation_data.price_money.amount / 100
          : "";

        // Retrieve the GTIN (UPC/EAN/ISBN)
        var gtin = variation.item_variation_data.upc || "";

        // Extract item option values for this variation
        var itemOptionValues = "";
        if (Array.isArray(variation.item_variation_data.item_option_values)) {
          itemOptionValues = JSON.stringify(variation.item_variation_data.item_option_values);
        }

        // New fields
        var sku = variation.item_variation_data.sku || "";
        var customAttributes = variation.custom_attribute_values ? JSON.stringify(variation.custom_attribute_values) : "";
        var measurementUnitId = variation.item_variation_data.measurement_unit_id || "";
        var pricingType = variation.item_variation_data.pricing_type || "";
        var updatedAt = variation.updated_at || item.updated_at || "";

        // Availability fields
        var presentAtAllLocations = variation.hasOwnProperty('present_at_all_locations') ? variation.present_at_all_locations : null;
        var presentAtLocationIds = Array.isArray(variation.present_at_location_ids) ? variation.present_at_location_ids : null;
        var absentAtLocationIds = Array.isArray(variation.absent_at_location_ids) ? variation.absent_at_location_ids : null;

        // If variation-level availability fields are null, use item-level fields
        if (presentAtAllLocations === null) {
          presentAtAllLocations = itemPresentAtAllLocations;
        }
        if (presentAtLocationIds === null) {
          presentAtLocationIds = itemPresentAtLocationIds;
        }
        if (absentAtLocationIds === null) {
          absentAtLocationIds = itemAbsentAtLocationIds;
        }

        // Add location-specific overrides
        var locationData = {};
        if (Array.isArray(variation.item_variation_data.location_overrides)) {
          variation.item_variation_data.location_overrides.forEach(function(override) {
            var locId = override.location_id;
            locationData[locId] = {
              track_inventory: override.track_inventory || false,
              inventory_alert_type: override.inventory_alert_type || "",
              inventory_alert_threshold: override.inventory_alert_threshold || ""
            };
          });
        }

        variationMap[variationId] = {
          variationId: variationId,
          itemId: itemId,
          itemName: itemName,
          description: description,
          itemUrl: itemUrl,
          variationName: variationName,
          price: price,
          gtin: gtin,
          isDeleted: isDeleted,
          catalogV1Ids: catalogV1Ids,
          presentAtAllLocations: presentAtAllLocations,
          presentAtLocationIds: presentAtLocationIds,
          absentAtLocationIds: absentAtLocationIds,
          itemVisibility: itemVisibility,
          categoryId: categoryId,
          categoryName: categoryName,
          modifierListInfo: modifierListInfo,
          productType: productType,
          skipModifierScreen: skipModifierScreen,
          taxIds: taxIds,
          itemOptions: itemOptions,
          itemOptionValues: itemOptionValues,
          sku: sku,
          customAttributes: customAttributes,
          measurementUnitId: measurementUnitId,
          pricingType: pricingType,
          availableOnline: availableOnline,
          availableForPickup: availableForPickup,
          updatedAt: updatedAt,
          locationData: locationData,
          images: [primaryImageUrl, secondaryImageUrl, tertiaryImageUrl]
        };
      });
    }
  });

  return variationMap;
}

// Function to fetch inventory counts for all variations using Retrieve Inventory Count endpoint
function fetchInventoryCountsForAllVariations(variationMap, locationIds, progressSheet) {
  var inventoryMap = {};
  var variationIds = Object.keys(variationMap);
  var headers = {
    "Square-Version": "2023-10-18",
    "Content-Type": "application/json"
  };

  var variationsProcessed = 0;

  for (var i = 0; i < variationIds.length; i++) {
    var variationId = variationIds[i];

    // Check if the user requested to stop processing
    var stopFlag = progressSheet.getRange('B5').getValue().toString().toUpperCase();
    if (stopFlag === 'STOP') {
      Logger.log('Processing halted by user during inventory fetching.');
      break;
    }

    // Throttle requests to avoid rate limits
    if (i > 0 && i % 100 == 0) {
      Logger.log('Sleeping for 1 second to avoid rate limits...');
      Utilities.sleep(1000); // Sleep for 1 second
    }

    var cursor = null;
    do {
      var url = 'https://connect.squareup.com/v2/inventory/' + variationId;
      if (locationIds.length > 0) {
        url += '?location_ids=' + locationIds.join(',');
      }
      if (cursor) {
        url += (locationIds.length > 0 ? '&' : '?') + 'cursor=' + cursor;
      }

      var options = {
        "method": "GET",
        "headers": headers,
        "muteHttpExceptions": true
      };

      var response = makeApiRequest(url, options);

      if (response.getResponseCode() === 200) {
        var data = JSON.parse(response.getContentText());
        if (Array.isArray(data.counts)) {
          data.counts.forEach(function(count) {
            var key = count.catalog_object_id + '_' + count.location_id;
            var quantity = parseInt(count.quantity || "0", 10);
            inventoryMap[key] = {
              quantity: quantity,
              availability: quantity > 0 ? 'Available' : 'Unavailable'
            };
          });
        } else {
          Logger.log("No counts found for variation ID " + variationId);
        }

        cursor = data.cursor || null;
      } else {
        Logger.log("Error retrieving inventory count for variation ID " + variationId + ": " + response.getContentText());
        displayAlert("Error retrieving inventory count for variation ID " + variationId + ": " + response.getContentText());
        break; // Exit the loop on error
      }
    } while (cursor);

    variationsProcessed++;

    // Update progress indicators every 100 variations
    if (variationsProcessed % 100 === 0) {
      var progressPercent = Math.round((variationsProcessed / variationIds.length) * 100);
      progressSheet.getRange('B2').setValue(variationsProcessed);
      progressSheet.getRange('B3').setValue(progressPercent);
      SpreadsheetApp.flush();
    }
  }

  Logger.log("Total Inventory Counts Retrieved: " + Object.keys(inventoryMap).length);
  return inventoryMap;
}

// Function to process variations and write data to the sheet
function processAndWriteData(sheet, variationMap, locationIds, inventoryMap, progressSheet) {
  // Prepare the header row dynamically to include all the extra fields and inventory columns for each location
  var headerRow = [
    "Variation ID (ID-B)", "Item ID (ID-A)", "Title", "Link", "Description", "Variation Name", "Price (CAD)",
    "GTIN (UPC/EAN/ISBN)", "SKU", "Custom Attributes", "Item Options", "Modifier Lists", "Product Type", "Measurement Unit",
    "Pricing Type", "Visibility", "Available Online", "Available for Pickup", "Updated At", "is_deleted",
    "catalog_v1_ids", "present_at_all_locations", "item_visibility", "category_id", "category_name",
    "modifier_list_info", "product_type", "skip_modifier_screen", "tax_ids", "item_option_values"
  ];

  // Add columns for location overrides for each location
  locationIds.forEach(function(locationId) {
    headerRow.push("Track Inventory at " + locationId);
    headerRow.push("Inventory Alert Type at " + locationId);
    headerRow.push("Inventory Alert Threshold at " + locationId);
  });

  // Add columns for availability at each location based on catalog data (active/inactive)
  locationIds.forEach(function(locationId) {
    headerRow.push("Is Active at " + locationId);
  });

  // Add columns for inventory counts at each location
  locationIds.forEach(function(locationId) {
    headerRow.push("Inventory at " + locationId);
  });

  headerRow.push("Image Link", "Additional Image Link 1", "Additional Image Link 2"); // Add extra image columns at the end

  // Write header row to the sheet
  sheet.appendRow(headerRow);

  var allRows = [];
  var variationsProcessed = 0;
  var stopProcessing = false; // Flag to control processing based on user input

  // Iterate over variationMap
  for (var variationId in variationMap) {
    if (variationMap.hasOwnProperty(variationId)) {
      // Check if the user requested to stop processing
      var stopFlag = progressSheet.getRange('B5').getValue().toString().toUpperCase();
      if (stopFlag === 'STOP') {
        Logger.log('Processing halted by user.');
        stopProcessing = true;
        break;
      }

      var variationData = variationMap[variationId];

      // Retrieve inventory counts for each location
      var inventoryCounts = [];
      var availabilityStatuses = [];

      locationIds.forEach(function(locationId) {
        var key = variationId + '_' + locationId;
        var inventoryInfo = inventoryMap.hasOwnProperty(key) ? inventoryMap[key] : { quantity: 0, availability: 'Unavailable' };
        inventoryCounts.push(inventoryInfo.quantity);
        availabilityStatuses.push(inventoryInfo.availability);
      });

      // Determine if variation is active at each location based on catalog data
      var activeStatuses = [];
      locationIds.forEach(function(locationId) {
        var isActive = isVariationActiveAtLocation(variationData, locationId);
        activeStatuses.push(isActive ? 'Active' : 'Inactive');
      });

      // Add location-specific overrides
      var locationOverrides = [];
      locationIds.forEach(function(locationId) {
        var locData = variationData.locationData[locationId] || {};
        locationOverrides.push(locData.track_inventory || "");
        locationOverrides.push(locData.inventory_alert_type || "");
        locationOverrides.push(locData.inventory_alert_threshold || "");
      });

      // Prepare the row data
      var rowData = [
        variationData.variationId, variationData.itemId, variationData.itemName, variationData.itemUrl,
        variationData.description, variationData.variationName, variationData.price, variationData.gtin,
        variationData.sku, variationData.customAttributes, variationData.itemOptions, variationData.modifierListInfo,
        variationData.productType, variationData.measurementUnitId, variationData.pricingType, variationData.itemVisibility,
        variationData.availableOnline, variationData.availableForPickup, variationData.updatedAt, variationData.isDeleted,
        variationData.catalogV1Ids, variationData.presentAtAllLocations, variationData.itemVisibility,
        variationData.categoryId, variationData.categoryName, variationData.modifierListInfo, variationData.productType,
        variationData.skipModifierScreen, variationData.taxIds, variationData.itemOptionValues
      ].concat(locationOverrides, activeStatuses, inventoryCounts, variationData.images);

      allRows.push(rowData);

      variationsProcessed++;

      // Update progress indicators every 100 variations
      if (variationsProcessed % 100 === 0) {
        var progressPercent = Math.round((variationsProcessed / Object.keys(variationMap).length) * 100);
        progressSheet.getRange('B2').setValue(variationsProcessed);
        progressSheet.getRange('B3').setValue(progressPercent);
        SpreadsheetApp.flush();
      }

      // Write data to the sheet in batches of 500 rows
      if (allRows.length >= 500) {
        var range = sheet.getRange(sheet.getLastRow() + 1, 1, allRows.length, headerRow.length);
        range.setValues(allRows);
        allRows = [];
      }
    }
  }

  // Write any remaining data to the sheet
  if (allRows.length > 0) {
    var range = sheet.getRange(sheet.getLastRow() + 1, 1, allRows.length, headerRow.length);
    range.setValues(allRows);
  }

  // Update final progress indicators
  progressSheet.getRange('B2').setValue(variationsProcessed);
  progressSheet.getRange('B3').setValue(100);
  SpreadsheetApp.flush();

  if (!stopProcessing) {
    displayAlert("Processing complete.");
  }
}

// Helper function to determine if a variation is active at a specific location
function isVariationActiveAtLocation(variationData, locationId) {
  if (variationData.presentAtAllLocations === true) {
    if (variationData.absentAtLocationIds && variationData.absentAtLocationIds.includes(locationId)) {
      return false; // Inactive at this location
    } else {
      return true; // Active at this location
    }
  } else {
    // presentAtAllLocations is false or undefined
    if (variationData.presentAtLocationIds && variationData.presentAtLocationIds.includes(locationId)) {
      return true; // Active at this location
    } else {
      return false; // Inactive at this location
    }
  }
}

// Function to get image URLs from image IDs and related objects
function getImageUrls(imageIds, relatedObjects) {
  if (Array.isArray(imageIds) && imageIds.length > 0) {
    return imageIds.map(function(imageId) {
      var imageObject = relatedObjects.find(function(obj) {
        return obj.id === imageId && obj.type === "IMAGE";
      });
      return imageObject ? imageObject.image_data.url : "";
    });
  } else {
    return [];
  }
}

// Fetch location IDs for the merchant and return as an array
function fetchLocationIds() {
  var locationApiUrl = 'https://connect.squareup.com/v2/locations';

  var headers = {
    "Square-Version": "2023-10-18",
    "Content-Type": "application/json"
  };

  var options = {
    "method": "GET",
    "headers": headers,
    "muteHttpExceptions": true
  };

  var response = makeApiRequest(locationApiUrl, options);
  var locationIds = [];

  if (response.getResponseCode() === 200) {
    var jsonData = JSON.parse(response.getContentText());
    if (Array.isArray(jsonData.locations) && jsonData.locations.length > 0) {
      locationIds = jsonData.locations.map(function(location) {
        return location.id;
      });
    } else {
      Logger.log("No locations found in the API response.");
      displayAlert("No locations found for this merchant.");
    }
  } else {
    Logger.log("Error retrieving locations: " + response.getContentText());
    displayAlert("Error retrieving locations. Check logs.");
  }

  return locationIds;
}

// Function to handle API requests and token management
function makeApiRequest(url, options) {
  var documentProperties = PropertiesService.getDocumentProperties();
  var accessToken = documentProperties.getProperty('SQUARE_ACCESS_TOKEN');

  if (!accessToken) {
    // Prompt the user to set the API key
    displayAlert('Access token is missing. Please use the "Set API Key" option in the "Square API" menu to provide your access token.');
    throw new Error('Access token is required to proceed. Please set it using the "Set API Key" menu option.');
  }

  // Ensure options.headers exists
  if (!options.headers) {
    options.headers = {};
  }

  // Ensure the Authorization header has the correct token
  options.headers["Authorization"] = "Bearer " + accessToken;

  var response = UrlFetchApp.fetch(url, options);
  var statusCode = response.getResponseCode();

  if (statusCode === 401) {
    // Unauthorized, token may be invalid or expired
    // Send email notification instead of prompting
    var emailAddress = documentProperties.getProperty('NOTIFICATION_EMAIL');

    if (emailAddress) {
      MailApp.sendEmail({
        to: emailAddress,
        subject: "Square Data Refresh Failed - Invalid Access Token",
        body: "The access token used for the Square API is invalid or expired. Please update it using the 'Set API Key' option in the 'Square API' menu."
      });
    } else {
      Logger.log('Notification email address is not set. Please use "Set Email Address" in the "Square API" menu.');
    }
    throw new Error('Access token is invalid or expired.');
  } else if (statusCode >= 200 && statusCode < 300) {
    // Success
    return response;
  } else {
    // Other errors
    Logger.log('API request failed with status code ' + statusCode + ': ' + response.getContentText());
    throw new Error('API request failed with status code ' + statusCode);
  }
}

// Helper function to display alerts only when UI is available
function displayAlert(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    // UI not available (e.g., running via trigger), so we skip showing the alert
    Logger.log("Alert: " + message);
  }
}
