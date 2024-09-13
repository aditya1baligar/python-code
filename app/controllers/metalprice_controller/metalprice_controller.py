from fastapi import FastAPI, HTTPException, Query, APIRouter
from pydantic import BaseModel
import requests

router = APIRouter()

base_api_url = "https://api.metalpriceapi.com/v1"
api_key = "5cf29824e60032bd26171b59706c24df"

# Sample CSV data
metal_codes = {
    "aluminum": "ALU",
    "gold": "XAU",
    "silver": "XAG",
    "copper": "XCU",
}

class MetalPriceResponse(BaseModel):
    formulated_value: float


@router.get("/get_formula_value", response_model=MetalPriceResponse)
async def get_formula_value(metal_name: str = "Aluminum", previous_date: str = Query(..., description="Date in the format YYYY-MM-DD")):
    metal_code = metal_codes.get(metal_name.lower())
    if not metal_code:
        raise HTTPException(status_code=404, detail="Metal name not found.")
    
    try:
        # Fetch current price
        response_current_price = requests.get(
            f"{base_api_url}/latest",
            params={"api_key": api_key, "base": "INR", "currencies": metal_code},
        )
        response_current_price.raise_for_status()  # Raise an error for bad status codes
        data_current = response_current_price.json()
        price_current = data_current["rates"].get(f"INR{metal_code}")
        
        if price_current is None:
            raise HTTPException(status_code=400, detail="Error fetching current price.")
        
        # Fetch previous price
        response_previous_price = requests.get(
            f"{base_api_url}/{previous_date}",
            params={"api_key": api_key, "base": "INR", "currencies": metal_code},
        )
        response_previous_price.raise_for_status()
        data_previous = response_previous_price.json()
        price_previous = data_previous["rates"].get(f"INR{metal_code}")
        
        if price_previous is None:
            raise HTTPException(status_code=400, detail="Error fetching previous price.")
        
        # Calculate the formulated value
        formulated_value = (price_previous - price_current) * 32.1507
        
        return MetalPriceResponse(formulated_value=round(formulated_value, 2))
    
    except requests.RequestException as e:
        raise HTTPException(status_code=500, detail=str(e))
    