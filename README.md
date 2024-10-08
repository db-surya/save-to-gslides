# Chrome Extension - Save to Google Slides

Chrome Extension that allows users to create presentations on Google Slides on the go. By simply selecting text or images from any webpage, you can create a new slide in Google Slides with just a few clicks. The selected text is automatically formatted as bullet points, and if the content exceeds the slide, it will automatically flow onto the next slide. This way, you don’t need to manually open Google Slides to make quick edits or additions.

## How to Set Up and Run the Project

Follow these steps to set up and run the project locally:

### 1. Clone the Repository  
Clone the project to your local machine using the following command:
```bash
git clone https://github.com/your-repo-url.git
```
### 2. Load the Extension in Chrome  
- Navigate to `chrome://extensions` in your Chrome browser.
- Enable **Developer Mode** (toggle in the top-right corner).
- Click **Load unpacked** and select the folder containing the cloned project.

### 3. Set Up OAuth 2.0 for Google Slides API  
To enable the extension to interact with Google Slides, follow these steps to set up OAuth 2.0:

#### 3.1 Create a Google API Project  
- Go to the [Google Cloud Console](https://console.cloud.google.com/).
- Create a new project.
- Enable the **Google Slides API** for your project.

#### 3.2 Set Up OAuth Consent Screen  
- Under **APIs & Services**, go to **OAuth consent screen**.
- Choose **External** for the user type and configure the necessary fields.
- Add your Chrome extension’s name, authorized domain, and other details.

#### 3.3 Create OAuth 2.0 Credentials  
- Navigate to **Credentials** and create an OAuth 2.0 Client ID.
- Choose **Application type: Web application**.
- Set the **Authorized redirect URIs** as follows:
- Copy the generated **Client ID** and **Client Secret** and replace it in the config.

![Alt Text for Image](https://drive.google.com/uc?export=view&id=1S3gFKYDHcOuhN686wWK5RQLbykK292jH)
