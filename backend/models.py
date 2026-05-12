from dataclasses import dataclass
from typing import Dict, List, Optional
from datetime import datetime

@dataclass
class ReachData:
    """Data extracted from REACH sheet"""
    actual_impressions: float
    link_clicks: float
    ctr: float
    reach: float
    frequency: float

@dataclass
class DeviceBreakdown:
    """Device performance data"""
    device_type: str
    impressions: float
    clicks: float
    ctr: float

@dataclass
class CreativeBreakdown:
    """Creative performance data"""
    name: str
    impressions: float
    clicks: float
    ctr: float

@dataclass
class AgeBreakdown:
    """Age band performance data"""
    age_band: str
    impressions: float
    clicks: float
    ctr: float

@dataclass
class GenderBreakdown:
    """Gender performance data"""
    gender: str
    impressions: float
    clicks: float
    ctr: float

@dataclass
class CampaignData:
    """Complete campaign data for one input file"""
    audience: str
    burst_number: str
    reach_data: ReachData
    device_breakdown: List[DeviceBreakdown]
    creative_breakdown: List[CreativeBreakdown]
    age_breakdown: List[AgeBreakdown]
    gender_breakdown: List[GenderBreakdown]
    start_date: Optional[datetime]
    end_date: Optional[datetime]

@dataclass
class TemplateMetadata:
    """Template metadata for output"""
    platform: str
    format_type: str
    booked_impressions: Dict[str, float]
    start_date: Dict[str, str]
    end_date: Dict[str, str]
    reporting_date: Dict[str, str]
    audience_labels: Dict[str, str]
