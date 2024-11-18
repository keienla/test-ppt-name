/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((_info) => { })

function openTaskpaneWithCode(event: Office.AddinCommands.Event) {
  Office.addin.showAsTaskpane()
}

function closeTaskpaneWithCode(event: Office.AddinCommands.Event) {
  Office.addin.hide()
}

// Register the function
Office.actions.associate("openTaskpaneWithCode", openTaskpaneWithCode);
Office.actions.associate("closeTaskpaneWithCode", closeTaskpaneWithCode);