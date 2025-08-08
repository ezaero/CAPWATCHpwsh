variable "wing_designator" {
  description = "Two-letter wing designation (e.g., CO, TX, CA)"
  type        = string
  default     = "CO"
  
  validation {
    condition     = can(regex("^[A-Z]{2}$", var.wing_designator))
    error_message = "Wing designator must be exactly 2 uppercase letters."
  }
}


variable "location" {
  description = "Azure region where resources will be created"
  type        = string
  default     = "West Central US"
}

variable "capwatch_org_id" {
  description = "CAPWATCH Organization ID for your wing"
  type        = string
  
  validation {
    condition     = can(regex("^[0-9]{1,4}$", var.capwatch_org_id))
    error_message = "CAPWATCH Organization ID must be 1-4 digits."
  }
}

variable "exchange_organization" {
  description = "Exchange Online organization domain (e.g., COCivilAirPatrol.onmicrosoft.com)"
  type        = string
  
  validation {
    condition     = can(regex(".*\\.onmicrosoft\\.com$", var.exchange_organization))
    error_message = "Exchange organization must end with .onmicrosoft.com."
  }
}

variable "timezone" {
  description = "Timezone for the Function App"
  type        = string
  default     = "Mountain Standard Time"
}

variable "log_email_to_address" {
  description = "Email address to receive log notifications (optional)"
  type        = string
  default     = ""
}

variable "log_email_from_address" {
  description = "Sender email address for notifications (optional)"
  type        = string
  default     = ""
}

variable "appinsights_connection_string" {
  description = "Application Insights connection string for logging and monitoring."
  type        = string
  default     = ""
}

