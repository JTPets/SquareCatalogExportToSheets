// Square API Google Sheets Script - Clean Production Version
// This script fetches Square catalog data and populates a Google Sheet

// This function creates a custom menu in the Google Sheets UI
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Square API')
    .addItem('Set API Key', 'setApiKey')
    .addItem('Set Email Address', 'setEmailAddress')
    .addItem('Set Store Domain', 'setStoreDomain')
    .addItem('Start Processing', 'startProcessing')
    .addItem('Set 3-Hour Timer', 'createDailyTrigger')
    .addSeparator()
    .addItem('🖼️ Populate Images from URLs', 'populateImagesFromUrls')
    .addItem('📂 Import Images from CSV', 'importImagesFromCsv')
    .addToUi();
}

// Function to prompt the user to enter their Square API key
function setApiKey() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Set Square API Key', 'Please enter your Square API access token:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    var apiKey = response.getResponseText().trim();

    if (apiKey) {
      var documentProperties = PropertiesService.getDocumentProperties();
      documentProperties.setProperty('SQUARE_ACCESS_TOKEN', apiKey);
      ui.alert('Success', 'Your Square API access token has been saved securely.', ui.ButtonSet.OK);
    } else {
      ui.alert('Error', 'No API key entered. Please try again.', ui.ButtonSet.OK);
    }
  } else {
    ui.alert('Operation cancelled.');
  }
}

// Function to prompt the user to enter their store domain
function setStoreDomain() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Set Store Domain', 'Please enter your Square Online Store domain (e.g., yourstore.com or yourstore.square.site):', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    var storeDomain = response.getResponseText().trim();

    if (storeDomain) {
      storeDomain = storeDomain.replace(/^https?:\/\//, '');
      var documentProperties = PropertiesService.getDocumentProperties();
      documentProperties.setProperty('STORE_DOMAIN', storeDomain);
      ui.alert('Success', 'Your store domain has been saved: ' + storeDomain, ui.ButtonSet.OK);
    } else {
      ui.alert('Error', 'No domain entered. Please try again.', ui.ButtonSet.OK);
    }
  } else {
    ui.alert('Operation cancelled.');
  }
}

function setEmailAddress() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Set Notification Email', 'Please enter your email address:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    var emailAddress = response.getResponseText().trim();

    if (emailAddress) {
      var documentProperties = PropertiesService.getDocumentProperties();
      documentProperties.setProperty('NOTIFICATION_EMAIL', emailAddress);
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
    // Get or create the API-Export sheet
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet();
      sheet.setName(sheetName);
    }

    // Setup Progress sheet
    var progressSheetName = 'Processing-Progress';
    var progressSheet = ss.getSheetByName(progressSheetName);
    if (progressSheet) {
      progressSheet.clear();
    } else {
      progressSheet = ss.insertSheet();
      progressSheet.setName(progressSheetName);
    }

    // Initialize progress indicators
    progressSheet.getRange('A1').setValue('Total Variations:');
    progressSheet.getRange('A2').setValue('Variations Processed:');
    progressSheet.getRange('A3').setValue('Progress (%):');
    progressSheet.getRange('A5').setValue('Type "STOP" in cell B5 to halt processing.');
    progressSheet.getRange('A6').setValue('Last Refreshed:');

    progressSheet.getRange('B5').setValue('');
    progressSheet.getRange('B6').setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));

    // Fetch all location IDs automatically
    var locationIds = fetchLocationIds();
    if (!locationIds.length) {
      Logger.log("No locations found for the merchant.");
      displayAlert("No locations found for this merchant.");
      return;
    }

    // Fetch all catalog data
    var catalogData = fetchAllCatalogData();
    if (catalogData.items.length === 0) {
      displayAlert("No items found in the catalog.");
      return;
    }

    progressSheet.getRange('B1').setValue(catalogData.variationCount);

    // Create variation map with reliable image handling
    var variationMap = buildVariationMapWithReliableImages(catalogData.items, catalogData.categoryMap, catalogData.imageMap);

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
    }

    Logger.log("Error: " + error.message);
    displayAlert("An error occurred: " + error.message);
  }
}

// Function to create a time-driven trigger to refresh data every 3 hours
function createDailyTrigger() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.prompt(
    'Set 3-Hour Refresh Schedule', 
    'At what hour (0-23) would you like the 3-hour refresh cycle to start?\nRefresh will run every 3 hours from this time.\nExample: 6 = 6AM, 9AM, 12PM, 3PM, 6PM, 9PM, 12AM, 3AM', 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() == ui.Button.OK) {
    var hourText = response.getResponseText().trim();
    var startHour = parseInt(hourText, 10);
    
    if (isNaN(startHour) || startHour < 0 || startHour > 23) {
      ui.alert('Error', 'Please enter a valid hour between 0 and 23.', ui.ButtonSet.OK);
      return;
    }
    
    deleteExistingTriggers();
    
    var hours = [];
    for (var i = 0; i < 24; i += 3) {
      var hour = (startHour + i) % 24;
      hours.push(hour);
      
      ScriptApp.newTrigger('startProcessing')
        .timeBased()
        .atHour(hour)
        .everyDays(1)
        .create();
    }
    
    var displayTimes = hours.map(function(hour) {
      return formatHourForDisplay(hour);
    }).join(', ');
    
    ui.alert('Success', 'Refresh schedule set for every 3 hours starting at ' + formatHourForDisplay(startHour) + '.\n\nRefresh times: ' + displayTimes, ui.ButtonSet.OK);
    Logger.log('3-hour triggers created starting at hour: ' + startHour + '. All hours: ' + hours.join(', '));
  } else {
    ui.alert('Operation cancelled.');
  }
}

function formatHourForDisplay(hour) {
  if (hour === 0) return '12:00 AM (Midnight)';
  if (hour === 12) return '12:00 PM (Noon)';
  if (hour < 12) return hour + ':00 AM';
  return (hour - 12) + ':00 PM';
}

function deleteExistingTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == 'startProcessing') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// Function to fetch all catalog data
function fetchAllCatalogData() {
  var allItems = [];
  var categoryMap = {};
  var imageMap = {};
  var variationCount = 0;
  var cursor = null;
  var listCatalogUrl = 'https://connect.squareup.com/v2/catalog/list';

  do {
    var response = fetchCatalogPage(listCatalogUrl, cursor);

    if (response.getResponseCode() === 200) {
      var jsonData = JSON.parse(response.getContentText());

      if (Array.isArray(jsonData.objects) && jsonData.objects.length > 0) {
        jsonData.objects.forEach(function(obj) {
          if (obj.type === 'ITEM') {
            allItems.push(obj);
            if (Array.isArray(obj.item_data.variations)) {
              variationCount += obj.item_data.variations.length;
            }
          } else if (obj.type === 'CATEGORY') {
            categoryMap[obj.id] = obj.category_data.name;
          } else if (obj.type === 'IMAGE') {
            if (obj.image_data && obj.image_data.url) {
              imageMap[obj.id] = {
                url: obj.image_data.url,
                name: obj.image_data.name || "",
                caption: obj.image_data.caption || ""
              };
            }
          }
        });
      }

      cursor = jsonData.cursor || null;
    } else {
      Logger.log("Error details from List Catalog: " + response.getContentText());
      displayAlert("Error retrieving catalog. Check logs for details.");
      return { items: [], categoryMap: {}, imageMap: {}, variationCount: 0 };
    }

  } while (cursor);

  Logger.log("Total Items: " + allItems.length);
  Logger.log("Total Categories: " + Object.keys(categoryMap).length);
  Logger.log("Total Images: " + Object.keys(imageMap).length);
  Logger.log("Total Variations: " + variationCount);

  return { 
    items: allItems, 
    categoryMap: categoryMap, 
    imageMap: imageMap, 
    variationCount: variationCount 
  };
}

function fetchCatalogPage(listCatalogUrl, cursor) {
  var headers = {
    "Square-Version": "2023-10-18",
    "Content-Type": "application/json"
  };

  var types = "ITEM,CATEGORY,IMAGE";
  var urlWithParams = listCatalogUrl + "?types=" + types + "&include_related_objects=true";
  
  if (cursor) {
    urlWithParams += "&cursor=" + cursor;
  }

  var listOptions = {
    "method": "GET",
    "headers": headers,
    "muteHttpExceptions": true
  };

  return makeApiRequest(urlWithParams, listOptions);
}

function createSlug(itemName) {
  if (!itemName) return "";
  
  return itemName.toLowerCase()
    .replace(/[àáâãäå]/g, 'a')
    .replace(/[èéêë]/g, 'e')
    .replace(/[ìíîï]/g, 'i')
    .replace(/[òóôõö]/g, 'o')
    .replace(/[ùúûü]/g, 'u')
    .replace(/[ñ]/g, 'n')
    .replace(/[ç]/g, 'c')
    .replace(/\s*&\s*/g, '-and-')
    .replace(/[''`"""]/g, '')
    .replace(/[^a-z0-9\s-]/g, '')
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-+|-+$/g, '');
}

function constructItemUrl(itemName, itemId, storeDomain) {
  if (!storeDomain || !itemId) return "";
  
  var slug = createSlug(itemName);
  if (!slug) return "";
  
  return "https://" + storeDomain + "/product/" + slug + "/" + itemId;
}

// Reliable image processing function
function buildVariationMapWithReliableImages(items, categoryMap, imageMap) {
  var variationMap = {};
  
  var documentProperties = PropertiesService.getDocumentProperties();
  var storeDomain = documentProperties.getProperty('STORE_DOMAIN');

  items.forEach(function(item) {
    if (item.item_data && Array.isArray(item.item_data.variations)) {
      var itemId = item.id || "";
      var itemName = item.item_data.name || "";
      var description = item.item_data.description || "";
      var itemUrl = constructItemUrl(itemName, itemId, storeDomain);

      // Reliable image processing - only use methods that work
      var imageUrls = [];
      
      // Method 1: Check if item has image_ids
      if (Array.isArray(item.image_ids) && item.image_ids.length > 0) {
        item.image_ids.forEach(function(imageId) {
          if (imageMap[imageId] && imageMap[imageId].url) {
            imageUrls.push(imageMap[imageId].url);
          }
        });
      }

      // Method 2: If no direct image_ids, try batch retrieve
      if (imageUrls.length === 0) {
        try {
          var fallbackImages = fetchItemWithImages(itemId);
          if (fallbackImages && fallbackImages.length > 0) {
            imageUrls = fallbackImages;
          }
        } catch (error) {
          Logger.log("Error in batch retrieve for " + itemId + ": " + error.message);
        }
      }

      var primaryImageUrl = (imageUrls.length > 0) ? imageUrls[0] : "";
      var secondaryImageUrl = (imageUrls.length > 1) ? imageUrls[1] : "";
      var tertiaryImageUrl = (imageUrls.length > 2) ? imageUrls[2] : "";

      // Process item data
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

      var itemPresentAtAllLocations = item.hasOwnProperty('present_at_all_locations') ? item.present_at_all_locations : false;
      var itemPresentAtLocationIds = Array.isArray(item.present_at_location_ids) ? item.present_at_location_ids : [];
      var itemAbsentAtLocationIds = Array.isArray(item.absent_at_location_ids) ? item.absent_at_location_ids : [];

      item.item_data.variations.forEach(function(variation) {
        var variationId = variation.id || "";
        var variationName = variation.item_variation_data.name || "";
        var price = variation.item_variation_data.price_money
          ? variation.item_variation_data.price_money.amount / 100
          : "";

        var gtin = variation.item_variation_data.upc || "";
        var itemOptionValues = "";
        if (Array.isArray(variation.item_variation_data.item_option_values)) {
          itemOptionValues = JSON.stringify(variation.item_variation_data.item_option_values);
        }

        var sku = variation.item_variation_data.sku || "";
        var customAttributes = variation.custom_attribute_values ? JSON.stringify(variation.custom_attribute_values) : "";
        var measurementUnitId = variation.item_variation_data.measurement_unit_id || "";
        var pricingType = variation.item_variation_data.pricing_type || "";
        var updatedAt = variation.updated_at || item.updated_at || "";

        var presentAtAllLocations = variation.hasOwnProperty('present_at_all_locations') ? variation.present_at_all_locations : null;
        var presentAtLocationIds = Array.isArray(variation.present_at_location_ids) ? variation.present_at_location_ids : null;
        var absentAtLocationIds = Array.isArray(variation.absent_at_location_ids) ? variation.absent_at_location_ids : null;

        if (presentAtAllLocations === null) {
          presentAtAllLocations = itemPresentAtAllLocations;
        }
        if (presentAtLocationIds === null) {
          presentAtLocationIds = itemPresentAtLocationIds;
        }
        if (absentAtLocationIds === null) {
          absentAtLocationIds = itemAbsentAtLocationIds;
        }

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

function fetchItemWithImages(itemId) {
  try {
    var url = 'https://connect.squareup.com/v2/catalog/batch-retrieve';
    var payload = {
      "object_ids": [itemId],
      "include_related_objects": true
    };
    
    var options = {
      "method": "POST",
      "headers": {
        "Square-Version": "2023-10-18",
        "Content-Type": "application/json"
      },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };
    
    var response = makeApiRequest(url, options);
    if (response.getResponseCode() === 200) {
      var data = JSON.parse(response.getContentText());
      
      var imageUrls = [];
      
      if (Array.isArray(data.related_objects)) {
        data.related_objects.forEach(function(obj) {
          if (obj.type === 'IMAGE' && obj.image_data && obj.image_data.url) {
            imageUrls.push(obj.image_data.url);
          }
        });
      }
      
      return imageUrls;
    }
  } catch (error) {
    Logger.log("Error in fetchItemWithImages: " + error.message);
  }
  return [];
}

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

    var stopFlag = progressSheet.getRange('B5').getValue().toString().toUpperCase();
    if (stopFlag === 'STOP') {
      Logger.log('Processing halted by user during inventory fetching.');
      break;
    }

    if (i > 0 && i % 50 == 0) {
      Logger.log('Sleeping for 2 seconds to avoid rate limits...');
      Utilities.sleep(2000);
    }

    var cursor = null;
    var retryCount = 0;
    var maxRetries = 3;
    
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
      var statusCode = response.getResponseCode();

      if (statusCode === 200) {
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
        }
        cursor = data.cursor || null;
        retryCount = 0;
      } else if (statusCode === 429 || statusCode >= 500) {
        retryCount++;
        if (retryCount <= maxRetries) {
          var waitTime = Math.pow(2, retryCount) * 1000;
          Logger.log("API error " + statusCode + " for variation " + variationId + ". Retrying in " + (waitTime/1000) + " seconds... (attempt " + retryCount + "/" + maxRetries + ")");
          Utilities.sleep(waitTime);
          continue;
        } else {
          Logger.log("Max retries exceeded for variation ID " + variationId + ". Skipping...");
          break;
        }
      } else {
        Logger.log("Error retrieving inventory count for variation ID " + variationId + ": " + response.getContentText());
        break;
      }
    } while (cursor && retryCount <= maxRetries);

    variationsProcessed++;

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

function processAndWriteData(sheet, variationMap, locationIds, inventoryMap, progressSheet) {
  sheet.clear();
  
  var headerRow = [
    "Variation ID (ID-B)", "Item ID (ID-A)", "Title", "Link", "Description", "Variation Name", "Price (CAD)",
    "GTIN (UPC/EAN/ISBN)", "SKU", "Custom Attributes", "Item Options", "Modifier Lists", "Product Type", "Measurement Unit",
    "Pricing Type", "Visibility", "Available Online", "Available for Pickup", "Updated At", "is_deleted",
    "catalog_v1_ids", "present_at_all_locations", "item_visibility", "category_id", "category_name",
    "modifier_list_info", "product_type", "skip_modifier_screen", "tax_ids", "item_option_values"
  ];

  locationIds.forEach(function(locationId) {
    headerRow.push("Track Inventory at " + locationId);
    headerRow.push("Inventory Alert Type at " + locationId);
    headerRow.push("Inventory Alert Threshold at " + locationId);
  });

  locationIds.forEach(function(locationId) {
    headerRow.push("Is Active at " + locationId);
  });

  locationIds.forEach(function(locationId) {
    headerRow.push("Inventory at " + locationId);
  });

  headerRow.push("Image Link", "Additional Image Link 1", "Additional Image Link 2");

  sheet.appendRow(headerRow);

  var allRows = [];
  var variationsProcessed = 0;
  var stopProcessing = false;

  for (var variationId in variationMap) {
    if (variationMap.hasOwnProperty(variationId)) {
      var stopFlag = progressSheet.getRange('B5').getValue().toString().toUpperCase();
      if (stopFlag === 'STOP') {
        Logger.log('Processing halted by user.');
        stopProcessing = true;
        break;
      }

      var variationData = variationMap[variationId];

      var inventoryCounts = [];
      var availabilityStatuses = [];

      locationIds.forEach(function(locationId) {
        var key = variationId + '_' + locationId;
        var inventoryInfo = inventoryMap.hasOwnProperty(key) ? inventoryMap[key] : { quantity: 0, availability: 'Unavailable' };
        inventoryCounts.push(inventoryInfo.quantity);
        availabilityStatuses.push(inventoryInfo.availability);
      });

      var activeStatuses = [];
      locationIds.forEach(function(locationId) {
        var isActive = isVariationActiveAtLocation(variationData, locationId);
        activeStatuses.push(isActive ? 'Active' : 'Inactive');
      });

      var locationOverrides = [];
      locationIds.forEach(function(locationId) {
        var locData = variationData.locationData[locationId] || {};
        locationOverrides.push(locData.track_inventory || "");
        locationOverrides.push(locData.inventory_alert_type || "");
        locationOverrides.push(locData.inventory_alert_threshold || "");
      });

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

      if (variationsProcessed % 100 === 0) {
        var progressPercent = Math.round((variationsProcessed / Object.keys(variationMap).length) * 100);
        progressSheet.getRange('B2').setValue(variationsProcessed);
        progressSheet.getRange('B3').setValue(progressPercent);
        SpreadsheetApp.flush();
      }

      if (allRows.length >= 500) {
        var range = sheet.getRange(sheet.getLastRow() + 1, 1, allRows.length, headerRow.length);
        range.setValues(allRows);
        allRows = [];
      }
    }
  }

  if (allRows.length > 0) {
    var range = sheet.getRange(sheet.getLastRow() + 1, 1, allRows.length, headerRow.length);
    range.setValues(allRows);
  }

  progressSheet.getRange('B2').setValue(variationsProcessed);
  progressSheet.getRange('B3').setValue(100);
  progressSheet.getRange('B6').setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
  SpreadsheetApp.flush();

  if (!stopProcessing) {
    displayAlert("Processing complete. If images are missing, use 'Populate Images from URLs' in the Square API menu.");
  }
}

function isVariationActiveAtLocation(variationData, locationId) {
  if (variationData.presentAtAllLocations === true) {
    if (variationData.absentAtLocationIds && variationData.absentAtLocationIds.includes(locationId)) {
      return false;
    } else {
      return true;
    }
  } else {
    if (variationData.presentAtLocationIds && variationData.presentAtLocationIds.includes(locationId)) {
      return true;
    } else {
      return false;
    }
  }
}

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

function makeApiRequest(url, options) {
  var documentProperties = PropertiesService.getDocumentProperties();
  var accessToken = documentProperties.getProperty('SQUARE_ACCESS_TOKEN');

  if (!accessToken) {
    displayAlert('Access token is missing. Please use the "Set API Key" option in the "Square API" menu to provide your access token.');
    throw new Error('Access token is required to proceed. Please set it using the "Set API Key" menu option.');
  }

  if (!options.headers) {
    options.headers = {};
  }

  options.headers["Authorization"] = "Bearer " + accessToken;

  var response = UrlFetchApp.fetch(url, options);
  var statusCode = response.getResponseCode();

  if (statusCode === 401) {
    var emailAddress = documentProperties.getProperty('NOTIFICATION_EMAIL');

    if (emailAddress) {
      MailApp.sendEmail({
        to: emailAddress,
        subject: "Square Data Refresh Failed - Invalid Access Token",
        body: "The access token used for the Square API is invalid or expired. Please update it using the 'Set API Key' option in the 'Square API' menu."
      });
    }
    throw new Error('Access token is invalid or expired.');
  } else if (statusCode >= 200 && statusCode < 300) {
    return response;
  } else {
    Logger.log('API request failed with status code ' + statusCode + ': ' + response.getContentText());
    throw new Error('API request failed with status code ' + statusCode);
  }
}

function displayAlert(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log("Alert: " + message);
  }
}

// ========================================
// BACKUP IMAGE SOLUTION - URL-Based Image Populator
// ========================================

function populateImagesFromUrls() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('API-Export');
  
  if (!sheet) {
    displayAlert("API-Export sheet not found. Run main processing first.");
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var linkColumnIndex = headers.indexOf('Link');
  var imageColumnIndex = headers.indexOf('Image Link');
  var additionalImage1Index = headers.indexOf('Additional Image Link 1');
  var additionalImage2Index = headers.indexOf('Additional Image Link 2');
  
  if (linkColumnIndex === -1) {
    displayAlert("Link column not found in the sheet.");
    return;
  }
  
  Logger.log("Starting image population from URLs...");
  var processedCount = 0;
  var foundImagesCount = 0;
  
  for (var i = 1; i < data.length; i++) {
    var productUrl = data[i][linkColumnIndex];
    var productTitle = data[i][2];
    
    if (!productUrl || data[i][imageColumnIndex]) {
      continue;
    }
    
    try {
      Logger.log("Processing: " + productTitle + " - " + productUrl);
      
      var response = UrlFetchApp.fetch(productUrl, {
        'muteHttpExceptions': true,
        'headers': {
          'User-Agent': 'Mozilla/5.0 (compatible; GoogleAppsScript)'
        }
      });
      
      if (response.getResponseCode() === 200) {
        var htmlContent = response.getContentText();
        var imageUrls = extractImageUrls(htmlContent, productUrl);
        
        if (imageUrls.length > 0) {
          sheet.getRange(i + 1, imageColumnIndex + 1).setValue(imageUrls[0]);
          foundImagesCount++;
        }
        if (imageUrls.length > 1) {
          sheet.getRange(i + 1, additionalImage1Index + 1).setValue(imageUrls[1]);
        }
        if (imageUrls.length > 2) {
          sheet.getRange(i + 1, additionalImage2Index + 1).setValue(imageUrls[2]);
        }
        
        processedCount++;
        Utilities.sleep(1000);
        
      } else {
        Logger.log("Failed to fetch " + productUrl + " - Status: " + response.getResponseCode());
      }
      
    } catch (error) {
      Logger.log("Error processing " + productUrl + ": " + error.message);
    }
    
    if (i % 10 === 0) {
      Logger.log("Processed " + processedCount + " items, found images for " + foundImagesCount + " products...");
      SpreadsheetApp.flush();
    }
  }
  
  Logger.log("Image population complete! Processed " + processedCount + " URLs, found images for " + foundImagesCount + " products.");
  displayAlert("Image population from URLs complete!\nProcessed: " + processedCount + " products\nFound images for: " + foundImagesCount + " products");
}

function extractImageUrls(htmlContent, baseUrl) {
  var imageUrls = [];
  
  var patterns = [
    /https:\/\/items-images-production\.s3\.us-west-2\.amazonaws\.com\/[^"'\s]+/g,
    /https:\/\/square-production\.s3\.amazonaws\.com\/[^"'\s]+/g,
    /<img[^>]+class="[^"]*product[^"]*"[^>]+src="([^"]+)"/gi,
    /<img[^>]+src="([^"]+)"[^>]+class="[^"]*product[^"]*"/gi,
    /<img[^>]+src="([^"]+\.(jpg|jpeg|png|webp))"[^>]*>/gi,
    /<meta[^>]+property="og:image"[^>]+content="([^"]+)"/gi,
    /"image"\s*:\s*"([^"]+)"/gi
  ];
  
  patterns.forEach(function(pattern) {
    var matches;
    while ((matches = pattern.exec(htmlContent)) !== null) {
      var imageUrl = matches[1] || matches[0];
      
      if (imageUrl.startsWith('//')) {
        imageUrl = 'https:' + imageUrl;
      } else if (imageUrl.startsWith('/')) {
        var domain = baseUrl.match(/https?:\/\/[^\/]+/);
        if (domain) {
          imageUrl = domain[0] + imageUrl;
        }
      }
      
      if (isValidProductImage(imageUrl)) {
        imageUrls.push(imageUrl);
      }
    }
  });
  
  var uniqueUrls = [];
  imageUrls.forEach(function(url) {
    if (uniqueUrls.indexOf(url) === -1) {
      uniqueUrls.push(url);
    }
  });
  
  return uniqueUrls.slice(0, 3);
}

function isValidProductImage(imageUrl) {
  if (!imageUrl || !imageUrl.match(/\.(jpg|jpeg|png|gif|webp)(\?|$)/i)) {
    return false;
  }
  
  var excludePatterns = [
    /logo/i, /favicon/i, /icon/i, /header/i, /footer/i, /banner/i,
    /cart/i, /checkout/i, /payment/i, /social/i, /placeholder/i,
    /default/i, /avatar/i, /profile/i
  ];
  
  for (var i = 0; i < excludePatterns.length; i++) {
    if (excludePatterns[i].test(imageUrl)) {
      return false;
    }
  }
  
  return true;
}

function importImagesFromCsv() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.prompt(
    'Import Images from CSV',
    'Please paste CSV content with columns: Product Name, Image URL',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  var csvContent = response.getResponseText().trim();
  
  if (!csvContent) {
    ui.alert('No CSV content provided.');
    return;
  }
  
  try {
    var lines = csvContent.split('\n');
    var imageMap = {};
    
    for (var i = 1; i < lines.length; i++) {
      var parts = lines[i].split(',');
      if (parts.length >= 2) {
        var productName = parts[0].trim().replace(/"/g, '');
        var imageUrl = parts[1].trim().replace(/"/g, '');
        imageMap[productName.toLowerCase()] = imageUrl;
      }
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('API-Export');
    
    if (!sheet) {
      ui.alert('API-Export sheet not found.');
      return;
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var titleIndex = headers.indexOf('Title');
    var imageIndex = headers.indexOf('Image Link');
    
    var matchCount = 0;
    
    for (var i = 1; i < data.length; i++) {
      var productTitle = data[i][titleIndex];
      if (productTitle && imageMap[productTitle.toLowerCase()]) {
        sheet.getRange(i + 1, imageIndex + 1).setValue(imageMap[productTitle.toLowerCase()]);
        matchCount++;
      }
    }
    
    ui.alert('Import complete! Matched ' + matchCount + ' products with images.');
    
  } catch (error) {
    ui.alert('Error processing CSV: ' + error.message);
  }
}
