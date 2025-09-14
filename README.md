# PowerPoint XR Sync Server
This Python server acts as a real-time bridge between a Microsoft PowerPoint presentation and an external client application (e.g., a Unity VR/AR environment). It monitors the presentation for state changes, sends updates to the client, and processes commands received from the client to modify the presentation.

## 🔧 Setup
### Prerequisites
* Windows Operating System
* Microsoft PowerPoint installed
* Python 3.8+

### Installation
1. Clone or download this repository.
2. Install the required Python packages using the `requirements.txt` file:
    ```
    pip install -r requirements.txt
    ```

### How to Run
1. Open a PowerPoint presentation.
2. Run the main server script from your terminal:
    ```
    python main.py
    ```
    _(Note: Assuming your main entry point is `main.py` which instantiates and starts `PowerPointServer`)_

The server will automatically connect to the active PowerPoint instance and begin listening for client connections.

## 📡 Communication Protocol
The server communicates over two ZeroMQ sockets on `localhost`.

### Server → Client (PUB Socket on `tcp://*:5557`)
The server broadcasts multipart messages to all connected clients.

* `CurrentViewRefs` (Optimized Slide Change)
  * Description: Sent when the slide changes. Contains only metadata for the views on the new slide.
  * Format: `[b"CurrentViewRefs", <slide_metadata_json>, <view1_metadata_json>, <view2_metadata_json>, ...]`

* `AnimationStep`
  * Description: Sent when an animation click occurs on the current slide.
  * Format: `[b"AnimationStep", <animation_metadata_json>]`
  * Example Payload: `{"slide": 1, "animation_step": 2}`

* `CurrentMode`
  * Description: Sent when the user switches between Edit and Presenter mode.
  * Format: `[b"CurrentMode", <mode_metadata_json>]`
  * Example Payload: `{"mode": "present"}`

* `AllViews` (On Request)
  * Description: Sends all views with full image data from the entire presentation. Used for initial client setup.
  * Format: `[b"AllViews", <view1_metadata_json>, <view1_image_bytes>, <view2_metadata_json>, <view2_image_bytes>, ...]`

### Client → Server (PULL Socket on `tcp://*:5558`)
The client sends commands to the server.

* `GetAllViews`
  * Description: Requests that the server broadcast all views using the `AllViews` message.
  * Format: `[b"GetAllViews"]`
* `CreateView`
  * Description: Asks the server to embed a new view into the current PowerPoint slide.
  * Format: `[b"CreateView", <view_metadata_json>, <view_thumbnail_image_bytes>]`