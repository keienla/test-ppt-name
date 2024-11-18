/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((_info) => { })

function openTaskpaneWithCode(event: Office.AddinCommands.Event) {
  Office.addin.showAsTaskpane()
}

// Register the function
Office.actions.associate("openTaskpaneWithCode", openTaskpaneWithCode);