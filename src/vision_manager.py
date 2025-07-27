import os
import requests
import base64
import json
from PIL import Image, UnidentifiedImageError
from io import BytesIO
import logging

logger = logging.getLogger('mcp_vision_manager')

# Theme definitions (copied from presentation_manager.py for consistency)
THEMES = {
    "modern_blue": {
        "background": (0x00, 0x5A, 0xC1),
        "text": (0xFF, 0xFF, 0xFF),
        "accent": (0xE0, 0xF7, 0xFA),
        "font": "Montserrat"
    },
    "elegant_green": {
        "background": (0x2E, 0x7D, 0x32),
        "text": (0xFF, 0xFF, 0xFF),
        "accent": (0xC8, 0xE6, 0xC9),
        "font": "Lato"
    }
}
class VisionManager:
    def __init__(self):
        """Initialize the VisionManager with Stable Diffusion configuration."""
        self.sd_url = os.environ.get('SD_WEBUI_URL', 'http://127.0.0.1:7860')
        self.auth_user = os.environ.get('SD_AUTH_USER')
        self.auth_pass = os.environ.get('SD_AUTH_PASS')
        # Load configuration for Stable Diffusion parameters
        self.config = self.load_config()
        logger.info(f"Initialized VisionManager with SD URL: {self.sd_url}")

    def load_config(self, config_path="config.json"):
         """Load configuration from config.json for Stable Diffusion parameters."""
         default_config = {
             "sd_params": {
                 "steps": 4,
                 "width": 1024,
                 "height": 1024,
                 "cfg_scale": 1,
                 "sampler_name": "Euler",
                 "seed": -1,
                 "n_iter": 1,
                 "scheduler": "Simple"
              }
         }
         try:
             if os.path.exists(config_path):
                with open(config_path, 'r') as f:
                    config = json.load(f)
                    logger.info(f"Loaded configuration from {config_path}")
                    return config

             logger.warning(f"Config file {config_path} not found, using default settings")

             return default_config
         except Exception as e:
             logger.error(f"Failed to load config: {str(e)}")
             return default_config

    async def generate_and_save_image(self, prompt: str, output_path: str, theme: str = "modern_blue") -> str:
        """Generate an image using Stable Diffusion API and save it with theme-aware post-processing."""
        # headers = {'Content-Type': 'application/json'}
        # auth = None
        # if self.auth_user and self.auth_pass:
        #     auth = (self.auth_user, self.auth_pass)
        headers = {'Content-Type': 'application/json'}
        auth = (self.auth_user, self.auth_pass) if self.auth_user and self.auth_pass else None
        theme_data = THEMES.get(theme, THEMES["modern_blue"])

        # payload = {
        #     "prompt": prompt,
        #     "negative_prompt": "",
        #     "steps": 4,
        #     "width": 1024,
        #     "height": 1024,
        #     "cfg_scale": 1,
        #     "sampler_name": "Euler",
        #     "seed": -1,
        #     "n_iter": 1,
        #     "scheduler": "Simple"

        # Append theme to prompt for consistent styling
        themed_prompt = f"{prompt}, high quality, {theme} tones"
        payload = {
            "prompt": themed_prompt,
            "negative_prompt": "",
            **self.config["sd_params"]
        }

        try:
            # Generate the image
            logger.info(f"Generating image with prompt: {themed_prompt}")
            response = requests.post(
                f"{self.sd_url}/sdapi/v1/txt2img",
                headers=headers,
                auth=auth,
                json=payload,
                timeout=3600
            )
            response.raise_for_status()
            
            if not response.json().get('images'):
                logger.error("No images generated")
                raise ValueError("No images generated")
            
            # Get the first image
            image_data = response.json()['images'][0]
            if ',' in image_data:
                image_data = image_data.split(',')[1]
            
            # Convert base64 to image
            image_bytes = base64.b64decode(image_data)
            image = Image.open(BytesIO(image_bytes)).convert("RGB")
            # Ensure the save directory exists
            try:
                image = image.resize((800, 600)).enhance(1.2)  # Resize and enhance contrast
                os.makedirs(os.path.dirname(output_path), exist_ok=True)
                image.save(output_path, format="PNG")
                logger.info(f"Saved processed image to {output_path}")
            except (IOError, OSError, UnidentifiedImageError) as e:
                logger.error(f"Failed to process/save image to {output_path}: {str(e)}")
                raise ValueError(f"Failed to process/save image to {output_path}: {str(e)}")
            #
            # # Save the image
            # image.save(output_path)
            return output_path
        except requests.RequestException as e:
            logger.error(f"Failed to generate image: {str(e)}")
            raise ValueError(f"Failed to generate image: {str(e)}")
        except Exception as e:
            logger.error(f"Unexpected error during image generation: {str(e)}")
            raise ValueError(f"Unexpected error during image generation: {str(e)}")
            
        # except requests.RequestException as e:
        #     raise ValueError(f"Failed to generate image: {str(e)}")
        # except (IOError, OSError) as e:
        #     raise ValueError(f"Failed to save image to {output_path}: {str(e)}")
        #
        # return output_path