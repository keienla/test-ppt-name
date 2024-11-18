# Test-ppt-name

This repo have as only purpose to show a bug in the addin Powerpoint. When open the taskpane using:

``` xml
<Action xsi:type="ExecuteFunction">
    <FunctionName>openTaskpaneWithCode</FunctionName>
</Action>
```

``` typescript
function openTaskpaneWithCode(event: Office.AddinCommands.Event) {
  Office.addin.showAsTaskpane()
}

// Register the function
Office.actions.associate("openTaskpaneWithCode", openTaskpaneWithCode);
```

The taskpane opened doesn't use the tag `<DisplayName DefaultValue="Title of Task Pane"/>` to show the title in the taskpane, and nothing is set.

---

Howerver, if the taskpane if first opened with an action as
``` xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>ButtonId1</TaskpaneId>
    <SourceLocation resid="Taskpane.Url"/>
</Action>
```

So in this case, the taskpane opened contain the value in the tag `DisplayName`