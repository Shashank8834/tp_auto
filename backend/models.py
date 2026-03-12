from pydantic import BaseModel, Field
from typing import Optional


class ConnectedPerson(BaseModel):
    name: str = Field(..., description="Name of the connected person")
    designation: str = Field(..., description="e.g. Managing Partner, Director")
    remuneration: float = Field(..., description="Annualized remuneration in AED")
    roles: str = Field(..., description="Roles and responsibilities description")


class RelatedParty(BaseModel):
    name: str = Field(default="", description="Name of the related party")
    relationship: str = Field(default="", description="Relationship with the company")
    nature_of_transaction: str = Field(default="", description="Nature & description of transaction")
    pricing_method: str = Field(default="", description="Pricing method used")


class CompanyInfo(BaseModel):
    company_name: str = Field(..., description="Full legal name of the company")
    company_short_name: str = Field(default="", description="Short name / alias for the company")
    nature_of_business: str = Field(..., description="Detailed description of business activities")
    address: str = Field(..., description="Address of the company")
    fiscal_year_start: str = Field(..., description="Fiscal year start date, e.g. '1st Feb 2024'")
    fiscal_year_end: str = Field(..., description="Fiscal year end date, e.g. '31st Jan 2025'")
    intangibles: str = Field(default="NA", description="Description of intangibles owned, or NA")
    activity_description: str = Field(default="", description="Short activity type, e.g. 'Event Management Services'")
    transaction_description: str = Field(default="", description="Description of controlled transactions, e.g. 'Event planning, venue management, catering'")


class TestedPartyFinancials(BaseModel):
    operating_revenue: float = Field(..., description="Total operating revenue in AED")
    cost_of_sales: float = Field(..., description="Cost of sales in AED")
    admin_expenses: float = Field(..., description="Admin & general expenses in AED")
    other_expenses: float = Field(..., description="Other expenses in AED")
    staff_salary: float = Field(..., description="Staff salary and benefits in AED")
    partner_salaries: float = Field(..., description="Partners salaries in AED")

    @property
    def total_operating_cost(self) -> float:
        return self.cost_of_sales + self.admin_expenses + self.other_expenses + self.staff_salary + self.partner_salaries

    @property
    def operating_profit(self) -> float:
        return self.operating_revenue - self.total_operating_cost

    @property
    def op_or_ratio(self) -> float:
        if self.operating_revenue == 0:
            return 0.0
        return self.operating_profit / self.operating_revenue

    @property
    def op_or_percentage(self) -> str:
        return f"{self.op_or_ratio * 100:.2f}%"


class GenerateReportRequest(BaseModel):
    company_info: CompanyInfo
    connected_persons: list[ConnectedPerson]
    related_parties: list[RelatedParty] = Field(default_factory=list)
    financials: TestedPartyFinancials
