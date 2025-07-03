# EMS Restart Service - UiPath RPA

This RPA project automates the process of remotely accessing the EMS (Enterprise Messaging System) server and restarting specific Windows services based on alerts received from **EMS inactive email notifications**. It uses **Computer Vision (CV)** to interact with the RDP interface, making it robust in environments where selectors are not available.

The process is **triggered automatically** by an **Outlook webhook** when a matching email arrives.

---

## Project Description

When an EMS inactive alert email is received, an Outlook-integrated webhook triggers this automation. The bot opens a Remote Desktop (RDP) session to the EMS server and uses **UiPath Computer Vision** to locate and restart specific Windows services. Email notifications are then sent to confirm whether the process succeeded or failed.

---

## Key Features

- Auto-triggered via **Outlook webhook**
- **Computer Vision-based automation** inside RDP
- Remote Windows service restart using `services.msc`
- Configurable services and credentials
- Success and failure email notification
- REFramework for reliability and modularity

---

## Project Structure

| Folder/File                  | Description                                                   |
|------------------------------|---------------------------------------------------------------|
| Main.xaml                    | Main entry point for the automation                          |
| Modular/RDP.xaml             | Opens and handles RDP session                                |
| Modular/EmailProcessing.xaml | Extracts and filters alert emails                            |
| Modular/EmailSuccess.xaml    | Sends success notification email                             |
| Modular/EmailError.xaml      | Sends error notification email                               |
| Framework/                   | REFramework components                                        |
| Data/Config.xlsx             | Configuration for services, server name, credentials         |
| project.json                 | UiPath project metadata                                       |
| README.md                    | Project documentation                                         |

---

## Process Workflow

### 1. **Trigger via Webhook**
- Email with subject like `EMS Inactive` is received
- Outlook webhook triggers bot execution

### 2. **Initialization**
- Read values from `Config.xlsx`:
  - Server credentials
  - Target services
  - Email parameters

### 3. **Remote Access (RDP)**
- `RDP.xaml` opens RDP session
- Waits for full remote desktop load

### 4. **Service Restart using CV**
- Launches `services.msc` inside RDP
- Uses **UiPath Computer Vision** to:
  - Locate service name
  - Right-click > Restart
- Repeat for all services listed in config

### 5. **Notification**
- On success: `EmailSuccess.xaml` sends confirmation  
- On failure: `EmailError.xaml` sends error report

### 6. **Logging & Error Handling**
- Take screenshots during error
- REFramework logs all steps
- Retry logic for recoverable steps (e.g., CV lag)

---

## How to Run

> Triggered via **Outlook webhook** integration (real-time)

Manual test:

1. Open in **UiPath Studio**
2. Configure `Config.xlsx`:
   - Email subject trigger
   - Server credentials
   - Services to restart
3. Simulate alert email or run manually
4. Optionally deploy to **Orchestrator** for monitoring/logging

---

## Exception Handling

- **Business Exceptions**:
  - No EMS email found
  - Service not visible via CV
- **System Exceptions**:
  - RDP disconnect
  - CV error (element not found, timeout)
- Logs:
  - Screenshots during error
  - Detailed exception message
- Email sent via `EmailError.xaml` for any failure

---

## Requirements

- UiPath Studio (Enterprise)
- Computer Vision package and API key (optional for cloud CV)
- RDP access to EMS server
- `services.msc` available on remote machine
- Outlook webhook or Power Automate integration
- Admin rights to restart services

---

## Contact

For questions, improvements, or collaboration:

- **Email:** fadillah650@gmail.com  
- **LinkedIn:** [Enrico Naufal Fadilla](https://linkedin.com/in/enrico-naufal-fadilla-54338a256)
