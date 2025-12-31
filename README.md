# SpatialSlides Python Middleware

SpatialSlides is a system that integrates immersive authoring directly into slide-based presentation workflows.

This Python server acts as the real-time bridge (or middleware) between the Microsoft PowerPoint COM API and the external Unity XR Client via ZeroMQ (ZMQ). Its core function is to monitor and transmit presentation state for immersive viewing and facilitate authoring while preserving the backward compatibility of the presentation file.

For the Unity Client source code and full system documentation, visit the main repository: **[SpatialSlides](https://github.com/JungWhoNam/SpatialSlides)**

-----

## Setup and Execution

### Prerequisites

  * **Operating System:** **Windows** (Required for Microsoft COM API).
  * **Software:** Microsoft PowerPoint (must be running with a presentation open).
  * **Python:** Python 3.x.

### Installation

1.  Clone or download this repository.
2.  Install the required Python packages using the **`requirements.txt`** file located in the repository root:
    ```bash
    pip install -r requirements.txt
    ```

### How to Run

1.  Open your PowerPoint presentation.
2.  Run the main server script from your terminal:
    ```bash
    python main.py
    ```
    *You can optionally pass custom ZMQ addresses or adjust the polling interval:*
    ```bash
    python main.py --pub_address tcp://*:5557 --rep_address tcp://*:5558 --interval 0.5
    ```
    **Note:** The `--rep_address` argument corresponds to the PULL socket binding.

-----

## Module Structure

The server's code is logically separated into two modules: `ppt_tools` (for PowerPoint I/O) and `ppt_server` (for server logic and coordination).

| Class | Module | Role & Responsibility                                                                                                                                     |
| :--- | :--- |:----------------------------------------------------------------------------------------------------------------------------------------------------------|
| **`PowerPointServer`** | `ppt_server` | The Orchestrator that runs the main polling loop, manages the lifecycle, and directs all communication. It aggregates the three primary components below. |
| **`PowerPointController`** | `ppt_tools` | Encapsulates all low-level communication with the Windows COM API (the only component that holds the live PowerPoint application reference).              |
| **`ZMQHandler`** | `ppt_server` | Manages the ZeroMQ sockets (PUB/SUB and PUSH/PULL) and the thread-safe message queue for network I/O.                                                     |
| **`SlideTracker`** | `ppt_server` | Focuses purely on state comparison logic (detecting a change in slide index, mode, or animation click) by polling the `PowerPointController`.            |

### System Constraints (Windows Only)

The server uses the **`win32com.client`** library to interface with PowerPoint's Component Object Model (COM) API. This makes the current server implementation limited to the Windows operating system.

### Metadata Storage

The spatial keyframe (3D view) metadata is stored by serializing the view transform into a JSON string and embedding it in the Alternative Text property of a small image shape placed outside the visible slide area. This technique preserves backward compatibility.

-----

## Communication Protocol

The server communicates over two ZeroMQ sockets on `localhost`, using a continuous polling interval (default: $1.0$ second) to check PowerPoint's state.

### Server $\to$ Client (PUB Socket on `tcp://*:5557`)

The server broadcasts multipart messages to all connected clients.

| Message | Trigger | Data Format | Description |
| :--- | :--- | :--- |:-------------------------------------------------------------------------------------------------------------------------|
| `CurrentViewRefs` | **Slide Change** or `GetCurrentViews` request | `[b"CurrentViewRefs", <slide_idx_json>, <meta1_json>, ...]` | Contains the JSON metadata for all spatial keyframes (3D views) linked to the new slide, enabling view synchronization. |
| `AnimationStep` | **Animation Click** on the current slide | `[b"AnimationStep", <anim_step_json>]` | Lightweight message indicating an animation or click step change. |
| `CurrentMode` | **Mode Change** (Edit $\leftrightarrow$ Presenter) | `[b"CurrentMode", <mode_json>]` | Notifies the client to switch between Authoring and Viewing modes. |
| `AllViews` | `GetAllViews` Request | `[b"AllViews", <meta1_json>, <img1_bytes>, ...]` | Sends all metadata and full image data from the entire presentation (used for initial authoring load). |

### Client $\to$ Server (PULL Socket on `tcp://*:5558`)

The XR client sends commands to the server, which processes them sequentially in a queue.

| Command | Format | Action Taken by Server |
| :--- | :--- |:-------------------------------------------------------------------------------------------------------------------------------------------------|
| `GetAllViews` | `[b"GetAllViews"]` | Requests that the server broadcast all views using the `AllViews` message. |
| `CreateView` | `[b"CreateView", <meta_json>, <img_bytes>]` | Executes the "snapshot" metaphor: Inserts the thumbnail image off-screen and embeds the spatial keyframe metadata into its Alternative Text. |
| `GetCurrentViews` | `[b"GetCurrentViews"]` | Requests a broadcast of the `CurrentViewRefs` message for the active slide. |
