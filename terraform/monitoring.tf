# Azure Monitor Action Group for email notifications
resource "azurerm_monitor_action_group" "capwatch_email" {
  name                = "capwatch-${local.resource_suffix}-alerts"
  resource_group_name = azurerm_resource_group.capwatch.name
  short_name          = "capwatchalerts"

  email_receiver {
    name                    = "log_email"
    email_address           = var.log_email_to_address
    use_common_alert_schema = true
  }
}

# Azure Monitor Metric Alert for Function App failures
resource "azurerm_monitor_metric_alert" "capwatch_func_failures" {
  name                = "capwatch-${local.resource_suffix}-func-failures"
  resource_group_name = azurerm_resource_group.capwatch.name
  scopes              = [azurerm_windows_function_app.capwatch.id]
  description         = "Alert on any function execution failure."
  severity            = 2
  enabled             = true
  frequency           = "PT5M"
  window_size         = "PT5M"

  criteria {
    metric_namespace = "Microsoft.Web/sites"
    metric_name      = "FunctionExecutionUnits"
    aggregation      = "Total"
    operator         = "GreaterThan"
    threshold        = 0
    dimension {
      name     = "Status"
      operator = "Include"
      values   = ["Failure"]
    }
  }

  action {
    action_group_id = azurerm_monitor_action_group.capwatch_email.id
  }

  depends_on = [azurerm_windows_function_app.capwatch]
}
