from flask import Flask, request, jsonify
from flask_cors import CORS
import openpyxl
import io

app = Flask(__name__)
CORS(app)


# calculate difference ratio
def calculate_difference_ratio(first, second):
    try:
        result = (first - second) / second
    except ZeroDivisionError:
        result = 0
    return result


# summation function
def summation(values):
    return sum(values)


# division function
def division(first, second):
    try:
        result = first / second
    except ZeroDivisionError:
        result = 0
    return result


# flask endpoint
@app.route("/upload", methods=["POST"])
def upload_file():
    # Initialization of variables

    # --- Start of CLI --- #
    # revenue
    cli_revenues_sales_of_real_estates = 0.0
    cli_revenues_rental = 0.0
    cli_revenues_management_fees = 0.0
    cli_revenues_hotel_operations = 0.0
    cli_total_revenue = 0.0

    # cos
    cli_cos_real_estates = 0.0
    cli_cos_depreciation = 0.0
    cli_cos_taxes = 0.0
    cli_cos_salaries_and_other_benefits = 0.0
    cli_cos_water = 0.0
    cli_cos_hotel = 0.0
    cli_total_cos = 0.0

    # gross profit
    cli_gp_sale_of_real_estates = 0.0
    cli_gp_rental = 0.0
    cli_gp_management_fees = 0.0
    cli_gp_water_income = 0.0
    cli_gp_hotel = 0.0
    cli_total_gross_profit = 0.0

    # operating expenses
    cli_operating_expenses_advertising = 0.0
    cli_operating_expenses_association_dues = 0.0
    cli_operating_expenses_commissions = 0.0
    cli_operating_expenses_communications = 0.0
    cli_operating_expenses_depreciation_and_amortization = 0.0
    cli_operating_expenses_donations = 0.0
    cli_operating_expenses_fuel_and_lubricants = 0.0
    cli_operating_expenses_impairment_losses = 0.0
    cli_operating_expenses_insurance = 0.0
    cli_operating_expenses_management_fee_expense = 0.0
    cli_operating_expenses_move_in_fees = 0.0
    cli_operating_expenses_penalties = 0.0
    cli_operating_expenses_professional_and_legal_fees = 0.0
    cli_operating_expenses_rent = 0.0
    cli_operating_expenses_repairs_and_maintenance = 0.0
    cli_operating_expenses_representation_and_entertainment = 0.0
    cli_operating_expenses_salaries_and_employee_benefits = 0.0
    cli_operating_expenses_security_and_janitorial_services = 0.0
    cli_operating_expenses_subscription_and_membership_dues = 0.0
    cli_operating_expenses_supplies = 0.0
    cli_operating_expenses_taxes_and_licenses = 0.0
    cli_operating_expenses_trainings_and_seminars = 0.0
    cli_operating_expenses_transportation_and_travel = 0.0
    cli_operating_expenses_utilities = 0.0
    cli_operating_expenses_other_operating_expenses = 0.0
    cli_total_operating_expenses = 0.0

    # operating income
    cli_reversal_of_payables = 0.0
    cli_administrative_charges = 0.0
    cli_reservation_fees_foregone = 0.0
    cli_late_payment_penalties = 0.0
    cli_documentation_fees = 0.0
    cli_sales_water = 0.0
    cli_water_income = 0.0
    cli_referal_incentives = 0.0
    cli_unrealized_foreign_exchange_gain = 0.0
    cli_realized_foreign_exchange_gain = 0.0
    cli_others = 0.0
    cli_other_operating_income = 0.0

    cli_net_operating_income = 0.0

    cli_equity_in_net_earnings_losses = 0.0
    cli_equity_in_net_earnings = 0.0
    cli_loss_on_sale_of_asset = 0.0
    cli_interest_expense_on_loans = 0.0
    cli_interest_expense_on_defined_benefit_obligation = 0.0
    cli_amortized_debt_issuance_cost = 0.0
    cli_expected_credit_losses = 0.0
    cli_interest_income_from_bank_deposits = 0.0
    cli_bank_charges = 0.0
    cli_interest_income_from_in_house_financing = 0.0
    cli_gain_on_sale_of_financial_assets = 0.0
    cli_other_gains = 0.0
    cli_gain_on_sale_of_property = 0.0
    cli_realized_foreign_exchange_loss = 0.0
    cli_unrealized_foreign_exchange_loss = 0.0
    cli_other_finance_income = 0.0
    cli_discount_on_non_current_contract_receivables = 0.0
    cli_other_finance_costs = 0.0
    cli_interest_expense_lease_liability = 0.0
    cli_total_other_income_or_expense = 0.0

    cli_net_profit_before_tax = 0.0

    cli_current_income_tax = 0.0
    cli_final_income_tax = 0.0
    cli_deferred_income_tax = 0.0
    cli_total_consolidated_net_income = 0.0

    cli_nci = 0.0
    cli_share_in_net_income = 0.0
    cli_total_nci = 0.0
    cli_total_net_income_attributable_to_parent = 0.0

    cli_consolidated_niat = 0.0
    cli_parent_niat = 0.0

    cli_gpm = 0.0
    cli_opex_ratio = 0.0
    cli_np_margin = 0.0
    # ---- End CLI ---- #

    # revenue
    blcbp_revenues_sales_of_real_estates = 0.0
    blcbp_revenues_rental = 0.0
    blcbp_revenues_management_fees = 0.0
    blcbp_revenues_hotel_operations = 0.0
    blcbp_total_revenue = 0.0

    # cos
    blcbp_cos_real_estates = 0.0
    blcbp_cos_depreciation = 0.0
    blcbp_cos_taxes = 0.0
    blcbp_cos_salaries_and_other_benefits = 0.0
    blcbp_cos_water = 0.0
    blcbp_cos_hotel = 0.0
    blcbp_total_cos = 0.0

    # gross profit
    blcbp_gp_sale_of_real_estates = 0.0
    blcbp_gp_rental = 0.0
    blcbp_gp_management_fees = 0.0
    blcbp_gp_hotel = 0.0
    blcbp_gp_water_income = 0.0
    blcbp_total_gross_profit = 0.0

    # operating expenses
    blcbp_operating_expenses_advertising = 0.0
    blcbp_operating_expenses_association_dues = 0.0
    blcbp_operating_expenses_commissions = 0.0
    blcbp_operating_expenses_communications = 0.0
    blcbp_operating_expenses_depreciation_and_amortization = 0.0
    blcbp_operating_expenses_donations = 0.0
    blcbp_operating_expenses_fuel_and_lubricants = 0.0
    blcbp_operating_expenses_impairment_losses = 0.0
    blcbp_operating_expenses_insurance = 0.0
    blcbp_operating_expenses_management_fee_expense = 0.0
    blcbp_operating_expenses_move_in_fees = 0.0
    blcbp_operating_expenses_penalties = 0.0
    blcbp_operating_expenses_professional_and_legal_fees = 0.0
    blcbp_operating_expenses_rent = 0.0
    blcbp_operating_expenses_repairs_and_maintenance = 0.0
    blcbp_operating_expenses_representation_and_entertainment = 0.0
    blcbp_operating_expenses_salaries_and_employee_benefits = 0.0
    blcbp_operating_expenses_security_and_janitorial_services = 0.0
    blcbp_operating_expenses_subscription_and_membership_dues = 0.0
    blcbp_operating_expenses_supplies = 0.0
    blcbp_operating_expenses_taxes_and_licenses = 0.0
    blcbp_operating_expenses_trainings_and_seminars = 0.0
    blcbp_operating_expenses_transportation_and_travel = 0.0
    blcbp_operating_expenses_utilities = 0.0
    blcbp_operating_expenses_other_operating_expenses = 0.0
    blcbp_total_operating_expenses = 0.0

    # operating income
    blcbp_reversal_of_payables = 0.0
    blcbp_administrative_charges = 0.0
    blcbp_reservation_fees_foregone = 0.0
    blcbp_late_payment_penalties = 0.0
    blcbp_documentation_fees = 0.0
    blcbp_sales_water = 0.0
    blcbp_water_income = 0.0
    blcbp_referal_incentives = 0.0
    blcbp_unrealized_foreign_exchange_gain = 0.0
    blcbp_realized_foreign_exchange_gain = 0.0
    blcbp_others = 0.0
    blcbp_other_operating_income = 0.0

    blcbp_net_operating_income = 0.0

    blcbp_equity_in_net_earnings_losses = 0.0
    blcbp_equity_in_net_earnings = 0.0
    blcbp_loss_on_sale_of_asset = 0.0
    blcbp_interest_expense_on_loans = 0.0
    blcbp_interest_expense_on_defined_benefit_obligation = 0.0
    blcbp_amortized_debt_issuance_cost = 0.0
    blcbp_expected_credit_losses = 0.0
    blcbp_interest_income_from_bank_deposits = 0.0
    blcbp_bank_charges = 0.0
    blcbp_interest_income_from_in_house_financing = 0.0
    blcbp_gain_on_sale_of_financial_assets = 0.0
    blcbp_other_gains = 0.0
    blcbp_gain_on_sale_of_property = 0.0
    blcbp_realized_foreign_exchange_loss = 0.0
    blcbp_unrealized_foreign_exchange_loss = 0.0
    blcbp_other_finance_income = 0.0
    blcbp_discount_on_non_current_contract_receivables = 0.0
    blcbp_other_finance_costs = 0.0
    blcbp_interest_expense_lease_liability = 0.0
    blcbp_total_other_income_or_expense = 0.0

    blcbp_net_profit_before_tax = 0.0

    blcbp_current_income_tax = 0.0
    blcbp_final_income_tax = 0.0
    blcbp_deferred_income_tax = 0.0
    blcbp_total_consolidated_net_income = 0.0

    blcbp_nci = 0.0
    blcbp_share_in_net_income = 0.0
    blcbp_total_nci = 0.0
    blcbp_total_net_income_attributable_to_parent = 0.0

    blcbp_consolidated_niat = 0.0
    blcbp_parent_niat = 0.0

    blcbp_gpm = 0.0
    blcbp_opex_ratio = 0.0
    blcbp_np_margin = 0.0
    # ---- End BLCBP ---- #

    # --- Start YES --- #

    # revenue
    yes_revenues_sales_of_real_estates = 0.0
    yes_revenues_rental = 0.0
    yes_revenues_management_fees = 0.0
    yes_revenues_hotel_operations = 0.0
    yes_total_revenue = 0.0

    # cos
    yes_cos_real_estates = 0.0
    yes_cos_depreciation = 0.0
    yes_cos_taxes = 0.0
    yes_cos_salaries_and_other_benefits = 0.0
    yes_cos_water = 0.0
    yes_cos_hotel = 0.0
    yes_total_cos = 0.0

    # gross profit
    yes_gp_sale_of_real_estates = 0.0
    yes_gp_rental = 0.0
    yes_gp_management_fees = 0.0
    yes_gp_hotel = 0.0
    yes_gp_water_income = 0.0
    yes_total_gross_profit = 0.0

    # operating expenses
    yes_operating_expenses_advertising = 0.0
    yes_operating_expenses_association_dues = 0.0
    yes_operating_expenses_commissions = 0.0
    yes_operating_expenses_communications = 0.0
    yes_operating_expenses_depreciation_and_amortization = 0.0
    yes_operating_expenses_donations = 0.0
    yes_operating_expenses_fuel_and_lubricants = 0.0
    yes_operating_expenses_impairment_losses = 0.0
    yes_operating_expenses_insurance = 0.0
    yes_operating_expenses_management_fee_expense = 0.0
    yes_operating_expenses_move_in_fees = 0.0
    yes_operating_expenses_penalties = 0.0
    yes_operating_expenses_professional_and_legal_fees = 0.0
    yes_operating_expenses_rent = 0.0
    yes_operating_expenses_repairs_and_maintenance = 0.0
    yes_operating_expenses_representation_and_entertainment = 0.0
    yes_operating_expenses_salaries_and_employee_benefits = 0.0
    yes_operating_expenses_security_and_janitorial_services = 0.0
    yes_operating_expenses_subscription_and_membership_dues = 0.0
    yes_operating_expenses_supplies = 0.0
    yes_operating_expenses_taxes_and_licenses = 0.0
    yes_operating_expenses_trainings_and_seminars = 0.0
    yes_operating_expenses_transportation_and_travel = 0.0
    yes_operating_expenses_utilities = 0.0
    yes_operating_expenses_other_operating_expenses = 0.0
    yes_total_operating_expenses = 0.0

    # operating income
    yes_reversal_of_payables = 0.0
    yes_administrative_charges = 0.0
    yes_reservation_fees_foregone = 0.0
    yes_late_payment_penalties = 0.0
    yes_documentation_fees = 0.0
    yes_sales_water = 0.0
    yes_water_income = 0.0
    yes_referal_incentives = 0.0
    yes_unrealized_foreign_exchange_gain = 0.0
    yes_realized_foreign_exchange_gain = 0.0
    yes_others = 0.0
    yes_other_operating_income = 0.0

    yes_net_operating_income = 0.0

    yes_equity_in_net_earnings_losses = 0.0
    yes_equity_in_net_earnings = 0.0
    yes_loss_on_sale_of_asset = 0.0
    yes_interest_expense_on_loans = 0.0
    yes_interest_expense_on_defined_benefit_obligation = 0.0
    yes_amortized_debt_issuance_cost = 0.0
    yes_expected_credit_losses = 0.0
    yes_interest_income_from_bank_deposits = 0.0
    yes_bank_charges = 0.0
    yes_interest_income_from_in_house_financing = 0.0
    yes_gain_on_sale_of_financial_assets = 0.0
    yes_other_gains = 0.0
    yes_gain_on_sale_of_property = 0.0
    yes_realized_foreign_exchange_loss = 0.0
    yes_unrealized_foreign_exchange_loss = 0.0
    yes_other_finance_income = 0.0
    yes_discount_on_non_current_contract_receivables = 0.0
    yes_other_finance_costs = 0.0
    yes_interest_expense_lease_liability = 0.0
    yes_total_other_income_or_expense = 0.0

    yes_net_profit_before_tax = 0.0

    yes_current_income_tax = 0.0
    yes_final_income_tax = 0.0
    yes_deferred_income_tax = 0.0
    yes_total_consolidated_net_income = 0.0

    yes_nci = 0.0
    yes_share_in_net_income = 0.0
    yes_total_nci = 0.0
    yes_total_net_income_attributable_to_parent = 0.0

    yes_consolidated_niat = 0.0
    yes_parent_niat = 0.0

    yes_gpm = 0.0
    yes_opex_ratio = 0.0
    yes_np_margin = 0.0
    # ---- End YES ---- #

    if "file" not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No selected file"}), 400
    if file and file.filename.endswith(".xlsx"):
        # Convert the file to a BytesIO object
        in_memory_file = io.BytesIO(file.read())

        # Load the workbook from the in-memory file
        workbook = openpyxl.load_workbook(in_memory_file)

        # Retrieve the year from the form data
        year = request.form.get("year")

        if year == "2020":
            sheet = workbook["Income Statement_2020"]

            # Extract the values from the specific cells

            # ---- CLI ---- #
            # revenue
            cli_revenues_sales_of_real_estates = sheet["C14"].value
            cli_revenues_rental = sheet["C15"].value
            cli_revenues_management_fees = sheet["C16"].value
            cli_revenues_hotel_operations = sheet["C17"].value
            cli_total_revenue = summation(
                [
                    cli_revenues_sales_of_real_estates,
                    cli_revenues_rental,
                    cli_revenues_management_fees,
                    cli_revenues_hotel_operations,
                ]
            )

            # cos
            cli_cos_real_estates = sheet["C19"].value
            cli_cos_depreciation = sheet["C20"].value
            cli_cos_taxes = sheet["C21"].value
            cli_cos_salaries_and_other_benefits = sheet["C22"].value
            cli_cos_water = 0.0
            cli_cos_hotel = sheet["C23"].value
            cli_total_cos = summation(
                [
                    cli_cos_real_estates,
                    cli_cos_depreciation,
                    cli_cos_taxes,
                    cli_cos_salaries_and_other_benefits,
                    cli_cos_water,
                    cli_cos_hotel,
                ]
            )

            # gross profit
            cli_gp_sale_of_real_estates = sheet["C25"].value
            cli_gp_rental = sheet["C26"].value
            cli_gp_management_fees = sheet["C27"].value
            cli_gp_hotel = sheet["C28"].value
            cli_total_gross_profit = summation(
                [
                    cli_gp_sale_of_real_estates,
                    cli_gp_rental,
                    cli_gp_management_fees,
                    cli_gp_hotel,
                ]
            )

            # operating expenses
            cli_operating_expenses_advertising = sheet["C31"].value
            cli_operating_expenses_association_dues = 0.0
            cli_operating_expenses_commissions = sheet["C32"].value
            cli_operating_expenses_communications = sheet["C33"].value
            cli_operating_expenses_depreciation_and_amortization = sheet["C34"].value
            cli_operating_expenses_donations = sheet["C35"].value
            cli_operating_expenses_fuel_and_lubricants = sheet["C36"].value
            cli_operating_expenses_impairment_losses = sheet["C37"].value
            cli_operating_expenses_insurance = sheet["C38"].value
            cli_operating_expenses_management_fee_expense = sheet["C39"].value
            cli_operating_expenses_move_in_fees = 0.0
            cli_operating_expenses_penalties = sheet["C40"].value
            cli_operating_expenses_professional_and_legal_fees = sheet["C41"].value
            cli_operating_expenses_rent = sheet["C42"].value
            cli_operating_expenses_repairs_and_maintenance = sheet["C43"].value
            cli_operating_expenses_representation_and_entertainment = sheet["C44"].value
            cli_operating_expenses_salaries_and_employee_benefits = sheet["C45"].value
            cli_operating_expenses_security_and_janitorial_services = sheet["C46"].value
            cli_operating_expenses_subscription_and_membership_dues = sheet["C47"].value
            cli_operating_expenses_supplies = sheet["C48"].value
            cli_operating_expenses_taxes_and_licenses = sheet["C49"].value
            cli_operating_expenses_trainings_and_seminars = sheet["C50"].value
            cli_operating_expenses_transportation_and_travel = sheet["C51"].value
            cli_operating_expenses_utilities = sheet["C52"].value
            cli_operating_expenses_other_operating_expenses = sheet["C53"].value
            cli_total_operating_expenses = summation(
                [
                    cli_operating_expenses_advertising,
                    cli_operating_expenses_commissions,
                    cli_operating_expenses_communications,
                    cli_operating_expenses_depreciation_and_amortization,
                    cli_operating_expenses_donations,
                    cli_operating_expenses_fuel_and_lubricants,
                    cli_operating_expenses_impairment_losses,
                    cli_operating_expenses_insurance,
                    cli_operating_expenses_management_fee_expense,
                    cli_operating_expenses_penalties,
                    cli_operating_expenses_professional_and_legal_fees,
                    cli_operating_expenses_rent,
                    cli_operating_expenses_repairs_and_maintenance,
                    cli_operating_expenses_salaries_and_employee_benefits,
                    cli_operating_expenses_security_and_janitorial_services,
                    cli_operating_expenses_subscription_and_membership_dues,
                    cli_operating_expenses_supplies,
                    cli_operating_expenses_taxes_and_licenses,
                    cli_operating_expenses_trainings_and_seminars,
                    cli_operating_expenses_transportation_and_travel,
                    cli_operating_expenses_utilities,
                    cli_operating_expenses_other_operating_expenses,
                ]
            )

            # operating income
            cli_reversal_of_payables = sheet["C56"].value
            cli_administrative_charges = sheet["C57"].value
            cli_reservation_fees_foregone = sheet["C58"].value
            cli_late_payment_penalties = sheet["C59"].value
            cli_documentation_fees = sheet["C60"].value
            cli_sales_water = sheet["C55"].value
            cli_water_income = sheet["C61"].value
            cli_referal_incentives = sheet["C62"].value
            cli_unrealized_foreign_exchange_gain = sheet["C63"].value
            cli_realized_foreign_exchange_gain = sheet["C64"].value
            cli_others = sheet["C65"].value

            cli_other_operating_income = summation(
                [
                    cli_sales_water,
                    cli_reversal_of_payables,
                    cli_administrative_charges,
                    cli_administrative_charges,
                    cli_reservation_fees_foregone,
                    cli_late_payment_penalties,
                    cli_documentation_fees,
                    cli_water_income,
                    cli_referal_incentives,
                    cli_unrealized_foreign_exchange_gain,
                    cli_realized_foreign_exchange_gain,
                    cli_others,
                ]
            )

            cli_net_operating_income = summation(
                [
                    cli_total_gross_profit,
                    cli_total_operating_expenses,
                    cli_other_operating_income,
                ]
            )

            cli_equity_in_net_earnings_losses = sheet["C68"].value
            cli_equity_in_net_earnings = sheet["C69"].value
            cli_loss_on_sale_of_asset = sheet["C70"].value
            cli_interest_expense_on_loans = sheet["C71"].value
            cli_interest_expense_on_defined_benefit_obligation = sheet["C72"].value
            cli_amortized_debt_issuance_cost = sheet["C73"].value
            cli_expected_credit_losses = sheet["C74"].value
            cli_interest_income_from_bank_deposits = sheet["C75"].value
            cli_bank_charges = sheet["C76"].value
            cli_interest_income_from_in_house_financing = sheet["C77"].value
            cli_gain_on_sale_of_financial_assets = sheet["C78"].value
            cli_other_gains = sheet["C79"].value
            cli_gain_on_sale_of_property = sheet["C80"].value
            cli_realized_foreign_exchange_loss = sheet["C81"].value
            cli_unrealized_foreign_exchange_loss = sheet["C82"].value
            cli_other_finance_income = sheet["C83"].value
            cli_discount_on_non_current_contract_receivables = sheet["C84"].value
            cli_other_finance_costs = sheet["C85"].value
            cli_interest_expense_lease_liability = sheet["C86"].value

            cli_total_other_income_or_expense = summation(
                [
                    cli_equity_in_net_earnings_losses,
                    cli_equity_in_net_earnings,
                    cli_loss_on_sale_of_asset,
                    cli_interest_expense_on_loans,
                    cli_interest_expense_on_defined_benefit_obligation,
                    cli_amortized_debt_issuance_cost,
                    cli_expected_credit_losses,
                    cli_interest_income_from_bank_deposits,
                    cli_bank_charges,
                    cli_interest_income_from_in_house_financing,
                    cli_gain_on_sale_of_financial_assets,
                    cli_other_gains,
                    cli_gain_on_sale_of_property,
                    cli_realized_foreign_exchange_loss,
                    cli_unrealized_foreign_exchange_loss,
                    cli_other_finance_income,
                    cli_discount_on_non_current_contract_receivables,
                    cli_other_finance_costs,
                    cli_interest_expense_lease_liability,
                ]
            )

            cli_net_profit_before_tax = (
                cli_net_operating_income + cli_total_other_income_or_expense
            )

            cli_current_income_tax = sheet["C89"].value
            cli_final_income_tax = sheet["C90"].value
            cli_deferred_income_tax = sheet["C91"].value

            cli_total_consolidated_net_income = (
                cli_net_profit_before_tax
                + cli_current_income_tax
                + cli_final_income_tax
                + cli_deferred_income_tax
            )

            cli_nci = sheet["C95"].value
            cli_share_in_net_income = sheet["C96"].value

            cli_total_nci = summation([cli_nci, cli_share_in_net_income])
            cli_total_net_income_attributable_to_parent = (
                cli_total_consolidated_net_income - cli_total_nci
            )

            cli_consolidated_niat = sheet["C106"].value
            cli_parent_niat = sheet["C107"].value

            cli_gpm = round((cli_total_gross_profit / cli_total_revenue) * 100)
            cli_opex_ratio = round(
                (cli_total_operating_expenses / cli_total_revenue) * 100
            )
            cli_np_margin = round(
                (cli_total_consolidated_net_income / cli_total_revenue) * 100
            )
            # ---- End CLI ---- #

            # ---- BLCBP ---- #
            # revenue
            blcbp_revenues_sales_of_real_estates = sheet["G14"].value
            blcbp_revenues_rental = sheet["G15"].value
            blcbp_revenues_management_fees = sheet["G16"].value
            blcbp_revenues_hotel_operations = sheet["G17"].value
            blcbp_total_revenue = summation(
                [
                    blcbp_revenues_sales_of_real_estates,
                    blcbp_revenues_rental,
                    blcbp_revenues_management_fees,
                    blcbp_revenues_hotel_operations,
                ]
            )

            # cos
            blcbp_cos_real_estates = sheet["G19"].value
            blcbp_cos_depreciation = sheet["G20"].value
            blcbp_cos_taxes = sheet["G21"].value
            blcbp_cos_salaries_and_other_benefits = sheet["G22"].value
            blcbp_cos_water = 0.0
            blcbp_cos_hotel = sheet["G23"].value
            blcbp_total_cos = summation(
                [
                    blcbp_cos_real_estates,
                    blcbp_cos_depreciation,
                    blcbp_cos_taxes,
                    blcbp_cos_salaries_and_other_benefits,
                    blcbp_cos_water,
                    blcbp_cos_hotel,
                ]
            )

            # gross profit
            blcbp_gp_sale_of_real_estates = sheet["G25"].value
            blcbp_gp_rental = sheet["G26"].value
            blcbp_gp_management_fees = sheet["G27"].value
            blcbp_gp_hotel = sheet["G28"].value
            blcbp_total_gross_profit = summation(
                [
                    blcbp_gp_sale_of_real_estates,
                    blcbp_gp_rental,
                    blcbp_gp_management_fees,
                    blcbp_gp_hotel,
                ]
            )

            # operating expenses
            blcbp_operating_expenses_advertising = sheet["G31"].value
            blcbp_operating_expenses_association_dues = 0.0
            blcbp_operating_expenses_commissions = sheet["G32"].value
            blcbp_operating_expenses_communications = sheet["G33"].value
            blcbp_operating_expenses_depreciation_and_amortization = sheet["G34"].value
            blcbp_operating_expenses_donations = sheet["G35"].value
            blcbp_operating_expenses_fuel_and_lubricants = sheet["G36"].value
            blcbp_operating_expenses_impairment_losses = sheet["G37"].value
            blcbp_operating_expenses_insurance = sheet["G38"].value
            blcbp_operating_expenses_management_fee_expense = sheet["G39"].value
            blcbp_operating_expenses_move_in_fees = 0.0
            blcbp_operating_expenses_penalties = sheet["G40"].value
            blcbp_operating_expenses_professional_and_legal_fees = sheet["G41"].value
            blcbp_operating_expenses_rent = sheet["G42"].value
            blcbp_operating_expenses_repairs_and_maintenance = sheet["G43"].value
            blcbp_operating_expenses_representation_and_entertainment = sheet[
                "G44"
            ].value
            blcbp_operating_expenses_salaries_and_employee_benefits = sheet["G45"].value
            blcbp_operating_expenses_security_and_janitorial_services = sheet[
                "G46"
            ].value
            blcbp_operating_expenses_subscription_and_membership_dues = sheet[
                "G47"
            ].value
            blcbp_operating_expenses_supplies = sheet["G48"].value
            blcbp_operating_expenses_taxes_and_licenses = sheet["G49"].value
            blcbp_operating_expenses_trainings_and_seminars = sheet["G50"].value
            blcbp_operating_expenses_transportation_and_travel = sheet["G51"].value
            blcbp_operating_expenses_utilities = sheet["G52"].value
            blcbp_operating_expenses_other_operating_expenses = sheet["G53"].value
            blcbp_total_operating_expenses = summation(
                [
                    blcbp_operating_expenses_advertising,
                    blcbp_operating_expenses_commissions,
                    blcbp_operating_expenses_communications,
                    blcbp_operating_expenses_depreciation_and_amortization,
                    blcbp_operating_expenses_donations,
                    blcbp_operating_expenses_fuel_and_lubricants,
                    blcbp_operating_expenses_impairment_losses,
                    blcbp_operating_expenses_insurance,
                    blcbp_operating_expenses_management_fee_expense,
                    blcbp_operating_expenses_penalties,
                    blcbp_operating_expenses_professional_and_legal_fees,
                    blcbp_operating_expenses_rent,
                    blcbp_operating_expenses_repairs_and_maintenance,
                    blcbp_operating_expenses_salaries_and_employee_benefits,
                    blcbp_operating_expenses_security_and_janitorial_services,
                    blcbp_operating_expenses_subscription_and_membership_dues,
                    blcbp_operating_expenses_supplies,
                    blcbp_operating_expenses_taxes_and_licenses,
                    blcbp_operating_expenses_trainings_and_seminars,
                    blcbp_operating_expenses_transportation_and_travel,
                    blcbp_operating_expenses_utilities,
                    blcbp_operating_expenses_other_operating_expenses,
                ]
            )

            # operating income
            blcbp_reversal_of_payables = sheet["G56"].value
            blcbp_administrative_charges = sheet["G57"].value
            blcbp_reservation_fees_foregone = sheet["G58"].value
            blcbp_late_payment_penalties = sheet["G59"].value
            blcbp_documentation_fees = sheet["G60"].value
            blcbp_sales_water = sheet["G55"].value
            blcbp_water_income = sheet["G61"].value
            blcbp_referal_incentives = sheet["G62"].value
            blcbp_unrealized_foreign_exchange_gain = sheet["G63"].value
            blcbp_realized_foreign_exchange_gain = sheet["G64"].value
            blcbp_others = sheet["G65"].value

            blcbp_other_operating_income = summation(
                [
                    blcbp_sales_water,
                    blcbp_reversal_of_payables,
                    blcbp_administrative_charges,
                    blcbp_administrative_charges,
                    blcbp_reservation_fees_foregone,
                    blcbp_late_payment_penalties,
                    blcbp_documentation_fees,
                    blcbp_water_income,
                    blcbp_referal_incentives,
                    blcbp_unrealized_foreign_exchange_gain,
                    blcbp_realized_foreign_exchange_gain,
                    blcbp_others,
                ]
            )

            blcbp_net_operating_income = summation(
                [
                    blcbp_total_gross_profit,
                    blcbp_total_operating_expenses,
                    blcbp_other_operating_income,
                ]
            )

            blcbp_equity_in_net_earnings_losses = sheet["G68"].value
            blcbp_equity_in_net_earnings = sheet["G69"].value
            blcbp_loss_on_sale_of_asset = sheet["G70"].value
            blcbp_interest_expense_on_loans = sheet["G71"].value
            blcbp_interest_expense_on_defined_benefit_obligation = sheet["G72"].value
            blcbp_amortized_debt_issuance_cost = sheet["G73"].value
            blcbp_expected_credit_losses = sheet["G74"].value
            blcbp_interest_income_from_bank_deposits = sheet["G75"].value
            blcbp_bank_charges = sheet["G76"].value
            blcbp_interest_income_from_in_house_financing = sheet["G77"].value
            blcbp_gain_on_sale_of_financial_assets = sheet["G78"].value
            blcbp_other_gains = sheet["G79"].value
            blcbp_gain_on_sale_of_property = sheet["G80"].value
            blcbp_realized_foreign_exchange_loss = sheet["G81"].value
            blcbp_unrealized_foreign_exchange_loss = sheet["G82"].value
            blcbp_other_finance_income = sheet["G83"].value
            blcbp_discount_on_non_current_contract_receivables = sheet["G84"].value
            blcbp_other_finance_costs = sheet["G85"].value
            blcbp_interest_expense_lease_liability = sheet["G86"].value

            blcbp_total_other_income_or_expense = summation(
                [
                    blcbp_equity_in_net_earnings_losses,
                    blcbp_equity_in_net_earnings,
                    blcbp_loss_on_sale_of_asset,
                    blcbp_interest_expense_on_loans,
                    blcbp_interest_expense_on_defined_benefit_obligation,
                    blcbp_amortized_debt_issuance_cost,
                    blcbp_expected_credit_losses,
                    blcbp_interest_income_from_bank_deposits,
                    blcbp_bank_charges,
                    blcbp_interest_income_from_in_house_financing,
                    blcbp_gain_on_sale_of_financial_assets,
                    blcbp_other_gains,
                    blcbp_gain_on_sale_of_property,
                    blcbp_realized_foreign_exchange_loss,
                    blcbp_unrealized_foreign_exchange_loss,
                    blcbp_other_finance_income,
                    blcbp_discount_on_non_current_contract_receivables,
                    blcbp_other_finance_costs,
                    blcbp_interest_expense_lease_liability,
                ]
            )

            blcbp_net_profit_before_tax = (
                blcbp_net_operating_income + blcbp_total_other_income_or_expense
            )

            blcbp_current_income_tax = sheet["G89"].value
            blcbp_final_income_tax = sheet["G90"].value
            blcbp_deferred_income_tax = sheet["G91"].value

            blcbp_total_consolidated_net_income = (
                blcbp_net_profit_before_tax
                + blcbp_current_income_tax
                + blcbp_final_income_tax
                + blcbp_deferred_income_tax
            )

            blcbp_nci = sheet["G95"].value
            blcbp_share_in_net_income = sheet["G96"].value

            blcbp_total_nci = summation([blcbp_nci, blcbp_share_in_net_income])
            blcbp_total_net_income_attributable_to_parent = (
                blcbp_total_consolidated_net_income - blcbp_total_nci
            )

            blcbp_consolidated_niat = sheet["G106"].value
            blcbp_parent_niat = sheet["G107"].value

            blcbp_gpm = round((blcbp_total_gross_profit / blcbp_total_revenue) * 100)
            blcbp_opex_ratio = round(
                (blcbp_total_operating_expenses / blcbp_total_revenue) * 100
            )
            blcbp_np_margin = round(
                (blcbp_total_consolidated_net_income / blcbp_total_revenue) * 100
            )
            # ---- End BLCBP ---- #

            # ---- YES ---- #
            # revenue
            yes_revenues_sales_of_real_estates = sheet["H14"].value
            yes_revenues_rental = sheet["H15"].value
            yes_revenues_management_fees = sheet["H16"].value
            yes_revenues_hotel_operations = sheet["H17"].value
            yes_total_revenue = summation(
                [
                    yes_revenues_sales_of_real_estates,
                    yes_revenues_rental,
                    yes_revenues_management_fees,
                    yes_revenues_hotel_operations,
                ]
            )

            # cos
            yes_cos_real_estates = sheet["H19"].value
            yes_cos_depreciation = sheet["H20"].value
            yes_cos_taxes = sheet["H21"].value
            yes_cos_salaries_and_other_benefits = sheet["H22"].value
            yes_cos_water = 0.0
            yes_cos_hotel = sheet["H23"].value
            yes_total_cos = summation(
                [
                    yes_cos_real_estates,
                    yes_cos_depreciation,
                    yes_cos_taxes,
                    yes_cos_salaries_and_other_benefits,
                    yes_cos_water,
                    yes_cos_hotel,
                ]
            )

            # gross profit
            yes_gp_sale_of_real_estates = sheet["H25"].value
            yes_gp_rental = sheet["H26"].value
            yes_gp_management_fees = sheet["H27"].value
            yes_gp_hotel = sheet["H28"].value
            yes_total_gross_profit = summation(
                [
                    yes_gp_sale_of_real_estates,
                    yes_gp_rental,
                    yes_gp_management_fees,
                    yes_gp_hotel,
                ]
            )

            # operating expenses
            yes_operating_expenses_advertising = sheet["H31"].value
            yes_operating_expenses_association_dues = 0.0
            yes_operating_expenses_commissions = sheet["H32"].value
            yes_operating_expenses_communications = sheet["H33"].value
            yes_operating_expenses_depreciation_and_amortization = sheet["H34"].value
            yes_operating_expenses_donations = sheet["H35"].value
            yes_operating_expenses_fuel_and_lubricants = sheet["H36"].value
            yes_operating_expenses_impairment_losses = sheet["H37"].value
            yes_operating_expenses_insurance = sheet["H38"].value
            yes_operating_expenses_management_fee_expense = sheet["H39"].value
            yes_operating_expenses_move_in_fees = 0.0
            yes_operating_expenses_penalties = sheet["H40"].value
            yes_operating_expenses_professional_and_legal_fees = sheet["H41"].value
            yes_operating_expenses_rent = sheet["H42"].value
            yes_operating_expenses_repairs_and_maintenance = sheet["H43"].value
            yes_operating_expenses_representation_and_entertainment = sheet["H44"].value
            yes_operating_expenses_salaries_and_employee_benefits = sheet["H45"].value
            yes_operating_expenses_security_and_janitorial_services = sheet["H46"].value
            yes_operating_expenses_subscription_and_membership_dues = sheet["H47"].value
            yes_operating_expenses_supplies = sheet["H48"].value
            yes_operating_expenses_taxes_and_licenses = sheet["H49"].value
            yes_operating_expenses_trainings_and_seminars = sheet["H50"].value
            yes_operating_expenses_transportation_and_travel = sheet["H51"].value
            yes_operating_expenses_utilities = sheet["H52"].value
            yes_operating_expenses_other_operating_expenses = sheet["H53"].value
            yes_total_operating_expenses = summation(
                [
                    yes_operating_expenses_advertising,
                    yes_operating_expenses_commissions,
                    yes_operating_expenses_communications,
                    yes_operating_expenses_depreciation_and_amortization,
                    yes_operating_expenses_donations,
                    yes_operating_expenses_fuel_and_lubricants,
                    yes_operating_expenses_impairment_losses,
                    yes_operating_expenses_insurance,
                    yes_operating_expenses_management_fee_expense,
                    yes_operating_expenses_penalties,
                    yes_operating_expenses_professional_and_legal_fees,
                    yes_operating_expenses_rent,
                    yes_operating_expenses_repairs_and_maintenance,
                    yes_operating_expenses_salaries_and_employee_benefits,
                    yes_operating_expenses_security_and_janitorial_services,
                    yes_operating_expenses_subscription_and_membership_dues,
                    yes_operating_expenses_supplies,
                    yes_operating_expenses_taxes_and_licenses,
                    yes_operating_expenses_trainings_and_seminars,
                    yes_operating_expenses_transportation_and_travel,
                    yes_operating_expenses_utilities,
                    yes_operating_expenses_other_operating_expenses,
                ]
            )

            # operating income
            yes_reversal_of_payables = sheet["H56"].value
            yes_administrative_charges = sheet["H57"].value
            yes_reservation_fees_foregone = sheet["H58"].value
            yes_late_payment_penalties = sheet["H59"].value
            yes_documentation_fees = sheet["H60"].value
            yes_sales_water = sheet["H55"].value
            yes_water_income = sheet["H61"].value
            yes_referal_incentives = sheet["H62"].value
            yes_unrealized_foreign_exchange_gain = sheet["H63"].value
            yes_realized_foreign_exchange_gain = sheet["H64"].value
            yes_others = sheet["H65"].value

            yes_other_operating_income = summation(
                [
                    yes_sales_water,
                    yes_reversal_of_payables,
                    yes_administrative_charges,
                    yes_administrative_charges,
                    yes_reservation_fees_foregone,
                    yes_late_payment_penalties,
                    yes_documentation_fees,
                    yes_water_income,
                    yes_referal_incentives,
                    yes_unrealized_foreign_exchange_gain,
                    yes_realized_foreign_exchange_gain,
                    yes_others,
                ]
            )

            yes_net_operating_income = summation(
                [
                    yes_total_gross_profit,
                    yes_total_operating_expenses,
                    yes_other_operating_income,
                ]
            )

            yes_equity_in_net_earnings_losses = sheet["H68"].value
            yes_equity_in_net_earnings = sheet["H69"].value
            yes_loss_on_sale_of_asset = sheet["H70"].value
            yes_interest_expense_on_loans = sheet["H71"].value
            yes_interest_expense_on_defined_benefit_obligation = sheet["H72"].value
            yes_amortized_debt_issuance_cost = sheet["H73"].value
            yes_expected_credit_losses = sheet["H74"].value
            yes_interest_income_from_bank_deposits = sheet["H75"].value
            yes_bank_charges = sheet["H76"].value
            yes_interest_income_from_in_house_financing = sheet["H77"].value
            yes_gain_on_sale_of_financial_assets = sheet["H78"].value
            yes_other_gains = sheet["H79"].value
            yes_gain_on_sale_of_property = sheet["H80"].value
            yes_realized_foreign_exchange_loss = sheet["H81"].value
            yes_unrealized_foreign_exchange_loss = sheet["H82"].value
            yes_other_finance_income = sheet["H83"].value
            yes_discount_on_non_current_contract_receivables = sheet["H84"].value
            yes_other_finance_costs = sheet["H85"].value
            yes_interest_expense_lease_liability = sheet["H86"].value

            yes_total_other_income_or_expense = summation(
                [
                    yes_equity_in_net_earnings_losses,
                    yes_equity_in_net_earnings,
                    yes_loss_on_sale_of_asset,
                    yes_interest_expense_on_loans,
                    yes_interest_expense_on_defined_benefit_obligation,
                    yes_amortized_debt_issuance_cost,
                    yes_expected_credit_losses,
                    yes_interest_income_from_bank_deposits,
                    yes_bank_charges,
                    yes_interest_income_from_in_house_financing,
                    yes_gain_on_sale_of_financial_assets,
                    yes_other_gains,
                    yes_gain_on_sale_of_property,
                    yes_realized_foreign_exchange_loss,
                    yes_unrealized_foreign_exchange_loss,
                    yes_other_finance_income,
                    yes_discount_on_non_current_contract_receivables,
                    yes_other_finance_costs,
                    yes_interest_expense_lease_liability,
                ]
            )

            yes_net_profit_before_tax = (
                yes_net_operating_income + yes_total_other_income_or_expense
            )

            yes_current_income_tax = sheet["H89"].value
            yes_final_income_tax = sheet["H90"].value
            yes_deferred_income_tax = sheet["H91"].value

            yes_total_consolidated_net_income = (
                yes_net_profit_before_tax
                + yes_current_income_tax
                + yes_final_income_tax
                + yes_deferred_income_tax
            )

            yes_nci = sheet["H95"].value
            yes_share_in_net_income = sheet["H96"].value

            yes_total_nci = summation([yes_nci, yes_share_in_net_income])
            yes_total_net_income_attributable_to_parent = (
                yes_total_consolidated_net_income - yes_total_nci
            )

            yes_consolidated_niat = sheet["H106"].value
            yes_parent_niat = sheet["H107"].value

            yes_gpm = round((yes_total_gross_profit / yes_total_revenue) * 100)
            yes_opex_ratio = round(
                (yes_total_operating_expenses / yes_total_revenue) * 100
            )
            yes_np_margin = round(
                (yes_total_consolidated_net_income / yes_total_revenue) * 100
            )
            # ---- End YES ---- #

        elif year == "2021":
            try:
                sheet = workbook["Income Statement_2021"]
            except:
                sheet = workbook["Income Stament_2021"]  # handle typo

            # Extract the values from the specific cells
            # ---- CLI ---- #
            # revenue
            cli_revenues_sales_of_real_estates = sheet["C14"].value
            cli_revenues_rental = sheet["C15"].value
            cli_revenues_management_fees = sheet["C16"].value
            cli_revenues_hotel_operations = sheet["C17"].value
            cli_total_revenue = summation(
                [
                    cli_revenues_sales_of_real_estates,
                    cli_revenues_rental,
                    cli_revenues_management_fees,
                    cli_revenues_hotel_operations,
                ]
            )

            # cos
            cli_cos_real_estates = sheet["C19"].value
            cli_cos_depreciation = sheet["C20"].value
            cli_cos_taxes = sheet["C21"].value
            cli_cos_salaries_and_other_benefits = sheet["C22"].value
            cli_cos_water = sheet["C23"].value
            cli_cos_hotel = sheet["C24"].value
            cli_total_cos = summation(
                [
                    cli_cos_real_estates,
                    cli_cos_depreciation,
                    cli_cos_taxes,
                    cli_cos_salaries_and_other_benefits,
                    cli_cos_water,
                    cli_cos_hotel,
                ]
            )

            # gross profit
            cli_gp_sale_of_real_estates = sheet["C26"].value
            cli_gp_rental = sheet["C27"].value
            cli_gp_management_fees = sheet["C28"].value
            cli_gp_water_income = sheet["C29"].value
            cli_gp_hotel = sheet["C30"].value
            cli_total_gross_profit = summation(
                [
                    cli_gp_sale_of_real_estates,
                    cli_gp_rental,
                    cli_gp_management_fees,
                    cli_gp_water_income,
                    cli_gp_hotel,
                ]
            )

            # operating expenses
            cli_operating_expenses_advertising = sheet["C33"].value
            cli_operating_expenses_association_dues = sheet["C34"].value
            cli_operating_expenses_commissions = sheet["C35"].value
            cli_operating_expenses_communications = sheet["C36"].value
            cli_operating_expenses_depreciation_and_amortization = sheet["C37"].value
            cli_operating_expenses_donations = sheet["C38"].value
            cli_operating_expenses_fuel_and_lubricants = sheet["C39"].value
            cli_operating_expenses_impairment_losses = sheet["C40"].value
            cli_operating_expenses_insurance = sheet["C41"].value
            cli_operating_expenses_management_fee_expense = sheet["C42"].value
            cli_operating_expenses_move_in_fees = sheet["C43"].value
            cli_operating_expenses_penalties = sheet["C44"].value
            cli_operating_expenses_professional_and_legal_fees = sheet["C45"].value
            cli_operating_expenses_rent = sheet["C46"].value
            cli_operating_expenses_repairs_and_maintenance = sheet["C47"].value
            cli_operating_expenses_representation_and_entertainment = sheet["C48"].value
            cli_operating_expenses_salaries_and_employee_benefits = sheet["C49"].value
            cli_operating_expenses_security_and_janitorial_services = sheet["C50"].value
            cli_operating_expenses_subscription_and_membership_dues = sheet["C51"].value
            cli_operating_expenses_supplies = sheet["C52"].value
            cli_operating_expenses_taxes_and_licenses = sheet["C53"].value
            cli_operating_expenses_trainings_and_seminars = sheet["C54"].value
            cli_operating_expenses_transportation_and_travel = sheet["C55"].value
            cli_operating_expenses_utilities = sheet["C56"].value
            cli_operating_expenses_other_operating_expenses = sheet["C57"].value
            cli_total_operating_expenses = summation(
                [
                    cli_operating_expenses_advertising,
                    cli_operating_expenses_commissions,
                    cli_operating_expenses_communications,
                    cli_operating_expenses_depreciation_and_amortization,
                    cli_operating_expenses_donations,
                    cli_operating_expenses_fuel_and_lubricants,
                    cli_operating_expenses_impairment_losses,
                    cli_operating_expenses_insurance,
                    cli_operating_expenses_management_fee_expense,
                    cli_operating_expenses_penalties,
                    cli_operating_expenses_professional_and_legal_fees,
                    cli_operating_expenses_rent,
                    cli_operating_expenses_repairs_and_maintenance,
                    cli_operating_expenses_salaries_and_employee_benefits,
                    cli_operating_expenses_security_and_janitorial_services,
                    cli_operating_expenses_subscription_and_membership_dues,
                    cli_operating_expenses_supplies,
                    cli_operating_expenses_taxes_and_licenses,
                    cli_operating_expenses_trainings_and_seminars,
                    cli_operating_expenses_transportation_and_travel,
                    cli_operating_expenses_utilities,
                    cli_operating_expenses_other_operating_expenses,
                ]
            )

            # operating income
            cli_reversal_of_payables = sheet["C60"].value
            cli_administrative_charges = sheet["C61"].value
            cli_reservation_fees_foregone = sheet["C62"].value
            cli_late_payment_penalties = sheet["C63"].value
            cli_documentation_fees = sheet["C64"].value
            cli_sales_water = sheet["C65"].value
            cli_water_income = sheet["C66"].value
            cli_referal_incentives = sheet["C67"].value
            cli_unrealized_foreign_exchange_gain = sheet["C68"].value
            cli_realized_foreign_exchange_gain = sheet["C69"].value
            cli_others = sheet["C70"].value

            cli_other_operating_income = summation(
                [
                    cli_sales_water,
                    cli_reversal_of_payables,
                    cli_administrative_charges,
                    cli_administrative_charges,
                    cli_reservation_fees_foregone,
                    cli_late_payment_penalties,
                    cli_documentation_fees,
                    cli_water_income,
                    cli_referal_incentives,
                    cli_unrealized_foreign_exchange_gain,
                    cli_realized_foreign_exchange_gain,
                    cli_others,
                ]
            )

            cli_net_operating_income = summation(
                [
                    cli_total_gross_profit,
                    cli_total_operating_expenses,
                    cli_other_operating_income,
                ]
            )

            cli_equity_in_net_earnings_losses = sheet["C73"].value
            cli_equity_in_net_earnings = sheet["C74"].value
            cli_loss_on_sale_of_asset = sheet["C75"].value
            cli_interest_expense_on_loans = sheet["C76"].value
            cli_interest_expense_on_defined_benefit_obligation = sheet["C77"].value
            cli_amortized_debt_issuance_cost = sheet["C78"].value
            cli_expected_credit_losses = sheet["C79"].value
            cli_interest_income_from_bank_deposits = sheet["C80"].value
            cli_bank_charges = sheet["C81"].value
            cli_interest_income_from_in_house_financing = sheet["C82"].value
            cli_gain_on_sale_of_financial_assets = sheet["C83"].value
            cli_other_gains = sheet["C84"].value
            cli_gain_on_sale_of_property = sheet["C85"].value
            cli_realized_foreign_exchange_loss = sheet["C86"].value
            cli_unrealized_foreign_exchange_loss = sheet["C87"].value
            cli_other_finance_income = sheet["C88"].value
            cli_discount_on_non_current_contract_receivables = sheet["C89"].value
            cli_other_finance_costs = sheet["C90"].value
            cli_interest_expense_lease_liability = sheet["C91"].value

            cli_total_other_income_or_expense = summation(
                [
                    cli_equity_in_net_earnings_losses,
                    cli_equity_in_net_earnings,
                    cli_loss_on_sale_of_asset,
                    cli_interest_expense_on_loans,
                    cli_interest_expense_on_defined_benefit_obligation,
                    cli_amortized_debt_issuance_cost,
                    cli_expected_credit_losses,
                    cli_interest_income_from_bank_deposits,
                    cli_bank_charges,
                    cli_interest_income_from_in_house_financing,
                    cli_gain_on_sale_of_financial_assets,
                    cli_other_gains,
                    cli_gain_on_sale_of_property,
                    cli_realized_foreign_exchange_loss,
                    cli_unrealized_foreign_exchange_loss,
                    cli_other_finance_income,
                    cli_discount_on_non_current_contract_receivables,
                    cli_other_finance_costs,
                    cli_interest_expense_lease_liability,
                ]
            )

            cli_net_profit_before_tax = (
                cli_net_operating_income + cli_total_other_income_or_expense
            )

            cli_current_income_tax = sheet["C94"].value
            cli_final_income_tax = sheet["C95"].value
            cli_deferred_income_tax = sheet["C96"].value

            cli_total_consolidated_net_income = (
                cli_net_profit_before_tax
                + cli_current_income_tax
                + cli_final_income_tax
                + cli_deferred_income_tax
            )

            cli_nci = 0.0
            cli_share_in_net_income = 0.0

            cli_total_nci = cli_total_consolidated_net_income * (1 - 100)
            cli_total_net_income_attributable_to_parent = (
                cli_total_consolidated_net_income - cli_total_nci
            )

            cli_consolidated_niat = calculate_difference_ratio(
                cli_total_consolidated_net_income, 1507601026
            )
            cli_parent_niat = calculate_difference_ratio(
                cli_total_consolidated_net_income, 1507601026
            )

            cli_gpm = round((cli_total_gross_profit / cli_total_revenue) * 100)
            cli_opex_ratio = round(
                (cli_total_operating_expenses / cli_total_revenue) * 100
            )
            cli_np_margin = round(
                (cli_total_consolidated_net_income / cli_total_revenue) * 100
            )
            # ---- End CLI ---- #

            # ---- BLCBP ---- #
            # revenue
            blcbp_revenues_sales_of_real_estates = sheet["G14"].value
            blcbp_revenues_rental = sheet["G15"].value
            blcbp_revenues_management_fees = sheet["G16"].value
            blcbp_revenues_hotel_operations = sheet["G17"].value
            blcbp_total_revenue = summation(
                [
                    blcbp_revenues_sales_of_real_estates,
                    blcbp_revenues_rental,
                    blcbp_revenues_management_fees,
                    blcbp_revenues_hotel_operations,
                ]
            )

            # cos
            blcbp_cos_real_estates = sheet["G19"].value
            blcbp_cos_depreciation = sheet["G20"].value
            blcbp_cos_taxes = sheet["G21"].value
            blcbp_cos_salaries_and_other_benefits = sheet["G22"].value
            blcbp_cos_water = sheet["G23"].value
            blcbp_cos_hotel = sheet["G24"].value
            blcbp_total_cos = summation(
                [
                    blcbp_cos_real_estates,
                    blcbp_cos_depreciation,
                    blcbp_cos_taxes,
                    blcbp_cos_salaries_and_other_benefits,
                    blcbp_cos_water,
                    blcbp_cos_hotel,
                ]
            )

            # gross profit
            blcbp_gp_sale_of_real_estates = sheet["G26"].value
            blcbp_gp_rental = sheet["G27"].value
            blcbp_gp_management_fees = sheet["G28"].value
            blcbp_gp_water_income = sheet["G29"].value
            blcbp_gp_hotel = sheet["G30"].value
            blcbp_total_gross_profit = summation(
                [
                    blcbp_gp_sale_of_real_estates,
                    blcbp_gp_rental,
                    blcbp_gp_management_fees,
                    blcbp_gp_water_income,
                    blcbp_gp_hotel,
                ]
            )

            # operating expenses
            blcbp_operating_expenses_advertising = sheet["G33"].value
            blcbp_operating_expenses_association_dues = sheet["G34"].value
            blcbp_operating_expenses_commissions = sheet["G35"].value
            blcbp_operating_expenses_communications = sheet["G36"].value
            blcbp_operating_expenses_depreciation_and_amortization = sheet["G37"].value
            blcbp_operating_expenses_donations = sheet["G38"].value
            blcbp_operating_expenses_fuel_and_lubricants = sheet["G39"].value
            blcbp_operating_expenses_impairment_losses = sheet["G40"].value
            blcbp_operating_expenses_insurance = sheet["G41"].value
            blcbp_operating_expenses_management_fee_expense = sheet["G42"].value
            blcbp_operating_expenses_move_in_fees = sheet["G43"].value
            blcbp_operating_expenses_penalties = sheet["G44"].value
            blcbp_operating_expenses_professional_and_legal_fees = sheet["G45"].value
            blcbp_operating_expenses_rent = sheet["G46"].value
            blcbp_operating_expenses_repairs_and_maintenance = sheet["G47"].value
            blcbp_operating_expenses_representation_and_entertainment = sheet[
                "G48"
            ].value
            blcbp_operating_expenses_salaries_and_employee_benefits = sheet["G49"].value
            blcbp_operating_expenses_security_and_janitorial_services = sheet[
                "G50"
            ].value
            blcbp_operating_expenses_subscription_and_membership_dues = sheet[
                "G51"
            ].value
            blcbp_operating_expenses_supplies = sheet["G52"].value
            blcbp_operating_expenses_taxes_and_licenses = sheet["G53"].value
            blcbp_operating_expenses_trainings_and_seminars = sheet["G54"].value
            blcbp_operating_expenses_transportation_and_travel = sheet["G55"].value
            blcbp_operating_expenses_utilities = sheet["G56"].value
            blcbp_operating_expenses_other_operating_expenses = sheet["G57"].value
            blcbp_total_operating_expenses = summation(
                [
                    blcbp_operating_expenses_advertising,
                    blcbp_operating_expenses_commissions,
                    blcbp_operating_expenses_communications,
                    blcbp_operating_expenses_depreciation_and_amortization,
                    blcbp_operating_expenses_donations,
                    blcbp_operating_expenses_fuel_and_lubricants,
                    blcbp_operating_expenses_impairment_losses,
                    blcbp_operating_expenses_insurance,
                    blcbp_operating_expenses_management_fee_expense,
                    blcbp_operating_expenses_penalties,
                    blcbp_operating_expenses_professional_and_legal_fees,
                    blcbp_operating_expenses_rent,
                    blcbp_operating_expenses_repairs_and_maintenance,
                    blcbp_operating_expenses_salaries_and_employee_benefits,
                    blcbp_operating_expenses_security_and_janitorial_services,
                    blcbp_operating_expenses_subscription_and_membership_dues,
                    blcbp_operating_expenses_supplies,
                    blcbp_operating_expenses_taxes_and_licenses,
                    blcbp_operating_expenses_trainings_and_seminars,
                    blcbp_operating_expenses_transportation_and_travel,
                    blcbp_operating_expenses_utilities,
                    blcbp_operating_expenses_other_operating_expenses,
                ]
            )

            # operating income
            blcbp_reversal_of_payables = sheet["G60"].value
            blcbp_administrative_charges = sheet["G61"].value
            blcbp_reservation_fees_foregone = sheet["G62"].value
            blcbp_late_payment_penalties = sheet["G63"].value
            blcbp_documentation_fees = sheet["G64"].value
            blcbp_sales_water = sheet["G65"].value
            blcbp_water_income = sheet["G66"].value
            blcbp_referal_incentives = sheet["G67"].value
            blcbp_unrealized_foreign_exchange_gain = sheet["G68"].value
            blcbp_realized_foreign_exchange_gain = sheet["G69"].value
            blcbp_others = sheet["G70"].value

            blcbp_other_operating_income = summation(
                [
                    blcbp_sales_water,
                    blcbp_reversal_of_payables,
                    blcbp_administrative_charges,
                    blcbp_administrative_charges,
                    blcbp_reservation_fees_foregone,
                    blcbp_late_payment_penalties,
                    blcbp_documentation_fees,
                    blcbp_water_income,
                    blcbp_referal_incentives,
                    blcbp_unrealized_foreign_exchange_gain,
                    blcbp_realized_foreign_exchange_gain,
                    blcbp_others,
                ]
            )

            blcbp_net_operating_income = summation(
                [
                    blcbp_total_gross_profit,
                    blcbp_total_operating_expenses,
                    blcbp_other_operating_income,
                ]
            )

            blcbp_equity_in_net_earnings_losses = sheet["G73"].value
            blcbp_equity_in_net_earnings = sheet["G74"].value
            blcbp_loss_on_sale_of_asset = sheet["G75"].value
            blcbp_interest_expense_on_loans = sheet["G76"].value
            blcbp_interest_expense_on_defined_benefit_obligation = sheet["G77"].value
            blcbp_amortized_debt_issuance_cost = sheet["G78"].value
            blcbp_expected_credit_losses = sheet["G79"].value
            blcbp_interest_income_from_bank_deposits = sheet["G80"].value
            blcbp_bank_charges = sheet["G81"].value
            blcbp_interest_income_from_in_house_financing = sheet["G82"].value
            blcbp_gain_on_sale_of_financial_assets = sheet["G83"].value
            blcbp_other_gains = sheet["G84"].value
            blcbp_gain_on_sale_of_property = sheet["G85"].value
            blcbp_realized_foreign_exchange_loss = sheet["G86"].value
            blcbp_unrealized_foreign_exchange_loss = sheet["G87"].value
            blcbp_other_finance_income = sheet["G88"].value
            blcbp_discount_on_non_current_contract_receivables = sheet["G89"].value
            blcbp_other_finance_costs = sheet["G90"].value
            blcbp_interest_expense_lease_liability = sheet["G91"].value

            blcbp_total_other_income_or_expense = summation(
                [
                    blcbp_equity_in_net_earnings_losses,
                    blcbp_equity_in_net_earnings,
                    blcbp_loss_on_sale_of_asset,
                    blcbp_interest_expense_on_loans,
                    blcbp_interest_expense_on_defined_benefit_obligation,
                    blcbp_amortized_debt_issuance_cost,
                    blcbp_expected_credit_losses,
                    blcbp_interest_income_from_bank_deposits,
                    blcbp_bank_charges,
                    blcbp_interest_income_from_in_house_financing,
                    blcbp_gain_on_sale_of_financial_assets,
                    blcbp_other_gains,
                    blcbp_gain_on_sale_of_property,
                    blcbp_realized_foreign_exchange_loss,
                    blcbp_unrealized_foreign_exchange_loss,
                    blcbp_other_finance_income,
                    blcbp_discount_on_non_current_contract_receivables,
                    blcbp_other_finance_costs,
                    blcbp_interest_expense_lease_liability,
                ]
            )

            blcbp_net_profit_before_tax = (
                blcbp_net_operating_income + blcbp_total_other_income_or_expense
            )

            blcbp_current_income_tax = sheet["G94"].value
            blcbp_final_income_tax = sheet["G95"].value
            blcbp_deferred_income_tax = sheet["G96"].value

            blcbp_total_consolidated_net_income = (
                blcbp_net_profit_before_tax
                + blcbp_current_income_tax
                + blcbp_final_income_tax
                + blcbp_deferred_income_tax
            )

            blcbp_nci = 0.0
            blcbp_share_in_net_income = 0.0

            blcbp_total_nci = blcbp_total_consolidated_net_income * (1 - 100)
            blcbp_total_net_income_attributable_to_parent = (
                blcbp_total_consolidated_net_income - blcbp_total_nci
            )

            blcbp_consolidated_niat = calculate_difference_ratio(
                blcbp_total_consolidated_net_income, -20178200
            )
            blcbp_parent_niat = calculate_difference_ratio(
                blcbp_total_consolidated_net_income, -20178200
            )

            blcbp_gpm = round((blcbp_total_gross_profit / blcbp_total_revenue) * 100)

            blcbp_opex_ratio = round(
                (blcbp_total_operating_expenses / blcbp_total_revenue) * 100
            )
            blcbp_np_margin = round(
                (blcbp_total_consolidated_net_income / blcbp_total_revenue) * 100
            )
            # ---- End BLCBP ---- #

            # ---- YES ---- #
            # revenue
            yes_revenues_sales_of_real_estates = sheet["H14"].value
            yes_revenues_rental = sheet["H15"].value
            yes_revenues_management_fees = sheet["H16"].value
            yes_revenues_hotel_operations = sheet["H17"].value
            yes_total_revenue = summation(
                [
                    yes_revenues_sales_of_real_estates,
                    yes_revenues_rental,
                    yes_revenues_management_fees,
                    yes_revenues_hotel_operations,
                ]
            )

            # cos
            yes_cos_real_estates = sheet["H19"].value
            yes_cos_depreciation = sheet["H20"].value
            yes_cos_taxes = sheet["H21"].value
            yes_cos_salaries_and_other_benefits = sheet["H22"].value
            yes_cos_water = sheet["H23"].value
            yes_cos_hotel = sheet["H24"].value
            yes_total_cos = summation(
                [
                    yes_cos_real_estates,
                    yes_cos_depreciation,
                    yes_cos_taxes,
                    yes_cos_salaries_and_other_benefits,
                    yes_cos_water,
                    yes_cos_hotel,
                ]
            )

            # gross profit
            yes_gp_sale_of_real_estates = sheet["H26"].value
            yes_gp_rental = sheet["H27"].value
            yes_gp_management_fees = sheet["H28"].value
            yes_gp_water_income = sheet["H29"].value
            yes_gp_hotel = sheet["H30"].value
            yes_total_gross_profit = summation(
                [
                    yes_gp_sale_of_real_estates,
                    yes_gp_rental,
                    yes_gp_management_fees,
                    yes_gp_water_income,
                    yes_gp_hotel,
                ]
            )

            # operating expenses
            yes_operating_expenses_advertising = sheet["H33"].value
            yes_operating_expenses_association_dues = sheet["H34"].value
            yes_operating_expenses_commissions = sheet["H35"].value
            yes_operating_expenses_communications = sheet["H36"].value
            yes_operating_expenses_depreciation_and_amortization = sheet["H37"].value
            yes_operating_expenses_donations = sheet["H38"].value
            yes_operating_expenses_fuel_and_lubricants = sheet["H39"].value
            yes_operating_expenses_impairment_losses = sheet["H40"].value
            yes_operating_expenses_insurance = sheet["H41"].value
            yes_operating_expenses_management_fee_expense = sheet["H42"].value
            yes_operating_expenses_move_in_fees = sheet["H43"].value
            yes_operating_expenses_penalties = sheet["H44"].value
            yes_operating_expenses_professional_and_legal_fees = sheet["H45"].value
            yes_operating_expenses_rent = sheet["H46"].value
            yes_operating_expenses_repairs_and_maintenance = sheet["H47"].value
            yes_operating_expenses_representation_and_entertainment = sheet["H48"].value
            yes_operating_expenses_salaries_and_employee_benefits = sheet["H49"].value
            yes_operating_expenses_security_and_janitorial_services = sheet["H50"].value
            yes_operating_expenses_subscription_and_membership_dues = sheet["H51"].value
            yes_operating_expenses_supplies = sheet["H52"].value
            yes_operating_expenses_taxes_and_licenses = sheet["H53"].value
            yes_operating_expenses_trainings_and_seminars = sheet["H54"].value
            yes_operating_expenses_transportation_and_travel = sheet["H55"].value
            yes_operating_expenses_utilities = sheet["H56"].value
            yes_operating_expenses_other_operating_expenses = sheet["H57"].value
            yes_total_operating_expenses = summation(
                [
                    yes_operating_expenses_advertising,
                    yes_operating_expenses_commissions,
                    yes_operating_expenses_communications,
                    yes_operating_expenses_depreciation_and_amortization,
                    yes_operating_expenses_donations,
                    yes_operating_expenses_fuel_and_lubricants,
                    yes_operating_expenses_impairment_losses,
                    yes_operating_expenses_insurance,
                    yes_operating_expenses_management_fee_expense,
                    yes_operating_expenses_penalties,
                    yes_operating_expenses_professional_and_legal_fees,
                    yes_operating_expenses_rent,
                    yes_operating_expenses_repairs_and_maintenance,
                    yes_operating_expenses_salaries_and_employee_benefits,
                    yes_operating_expenses_security_and_janitorial_services,
                    yes_operating_expenses_subscription_and_membership_dues,
                    yes_operating_expenses_supplies,
                    yes_operating_expenses_taxes_and_licenses,
                    yes_operating_expenses_trainings_and_seminars,
                    yes_operating_expenses_transportation_and_travel,
                    yes_operating_expenses_utilities,
                    yes_operating_expenses_other_operating_expenses,
                ]
            )

            # operating income
            yes_reversal_of_payables = sheet["H60"].value
            yes_administrative_charges = sheet["H61"].value
            yes_reservation_fees_foregone = sheet["H62"].value
            yes_late_payment_penalties = sheet["H63"].value
            yes_documentation_fees = sheet["H64"].value
            yes_sales_water = sheet["H65"].value
            yes_water_income = sheet["H66"].value
            yes_referal_incentives = sheet["H67"].value
            yes_unrealized_foreign_exchange_gain = sheet["H68"].value
            yes_realized_foreign_exchange_gain = sheet["H69"].value
            yes_others = sheet["H70"].value

            yes_other_operating_income = summation(
                [
                    yes_sales_water,
                    yes_reversal_of_payables,
                    yes_administrative_charges,
                    yes_administrative_charges,
                    yes_reservation_fees_foregone,
                    yes_late_payment_penalties,
                    yes_documentation_fees,
                    yes_water_income,
                    yes_referal_incentives,
                    yes_unrealized_foreign_exchange_gain,
                    yes_realized_foreign_exchange_gain,
                    yes_others,
                ]
            )

            yes_net_operating_income = summation(
                [
                    yes_total_gross_profit,
                    yes_total_operating_expenses,
                    yes_other_operating_income,
                ]
            )

            yes_equity_in_net_earnings_losses = sheet["H73"].value
            yes_equity_in_net_earnings = sheet["H74"].value
            yes_loss_on_sale_of_asset = sheet["H75"].value
            yes_interest_expense_on_loans = sheet["H76"].value
            yes_interest_expense_on_defined_benefit_obligation = sheet["H77"].value
            yes_amortized_debt_issuance_cost = sheet["H78"].value
            yes_expected_credit_losses = sheet["H79"].value
            yes_interest_income_from_bank_deposits = sheet["H80"].value
            yes_bank_charges = sheet["H81"].value
            yes_interest_income_from_in_house_financing = sheet["H82"].value
            yes_gain_on_sale_of_financial_assets = sheet["H83"].value
            yes_other_gains = sheet["H84"].value
            yes_gain_on_sale_of_property = sheet["H85"].value
            yes_realized_foreign_exchange_loss = sheet["H86"].value
            yes_unrealized_foreign_exchange_loss = sheet["H87"].value
            yes_other_finance_income = sheet["H88"].value
            yes_discount_on_non_current_contract_receivables = sheet["H89"].value
            yes_other_finance_costs = sheet["H90"].value
            yes_interest_expense_lease_liability = sheet["H91"].value

            yes_total_other_income_or_expense = summation(
                [
                    yes_equity_in_net_earnings_losses,
                    yes_equity_in_net_earnings,
                    yes_loss_on_sale_of_asset,
                    yes_interest_expense_on_loans,
                    yes_interest_expense_on_defined_benefit_obligation,
                    yes_amortized_debt_issuance_cost,
                    yes_expected_credit_losses,
                    yes_interest_income_from_bank_deposits,
                    yes_bank_charges,
                    yes_interest_income_from_in_house_financing,
                    yes_gain_on_sale_of_financial_assets,
                    yes_other_gains,
                    yes_gain_on_sale_of_property,
                    yes_realized_foreign_exchange_loss,
                    yes_unrealized_foreign_exchange_loss,
                    yes_other_finance_income,
                    yes_discount_on_non_current_contract_receivables,
                    yes_other_finance_costs,
                    yes_interest_expense_lease_liability,
                ]
            )

            yes_net_profit_before_tax = (
                yes_net_operating_income + yes_total_other_income_or_expense
            )

            yes_current_income_tax = sheet["H94"].value
            yes_final_income_tax = sheet["H95"].value
            yes_deferred_income_tax = sheet["H96"].value

            yes_total_consolidated_net_income = (
                yes_net_profit_before_tax
                + yes_current_income_tax
                + yes_final_income_tax
                + yes_deferred_income_tax
            )

            yes_nci = 0.0
            yes_share_in_net_income = 0.0

            yes_total_nci = yes_total_consolidated_net_income * (1 - 100)
            yes_total_net_income_attributable_to_parent = (
                yes_total_consolidated_net_income - yes_total_nci
            )

            yes_consolidated_niat = calculate_difference_ratio(
                yes_total_consolidated_net_income, 5949144
            )
            yes_parent_niat = calculate_difference_ratio(
                yes_total_consolidated_net_income, 5949144
            )

            yes_gpm = round((yes_total_gross_profit / yes_total_revenue) * 100)
            yes_opex_ratio = round(
                (yes_total_operating_expenses / yes_total_revenue) * 100
            )
            yes_np_margin = round(
                (yes_total_consolidated_net_income / yes_total_revenue) * 100
            )
            # ---- End YES ---- #

        else:
            return jsonify({"error": "Invalid year"}), 400

        # Example of JSON response with multiple values
        return (
            jsonify(
                {
                    "CLI": {
                        "revenues_sales_of_real_estates": cli_revenues_sales_of_real_estates,
                        "revenues_rental": cli_revenues_rental,
                        "revenues_management_fees": cli_revenues_management_fees,
                        "revenues_hotel_operations": cli_revenues_hotel_operations,
                        "total_revenue": cli_total_revenue,
                        "cos_real_estates": cli_cos_real_estates,
                        "cos_depreciation": cli_cos_depreciation,
                        "cos_taxes": cli_cos_taxes,
                        "cos_salaries_and_other_benefits": cli_cos_salaries_and_other_benefits,
                        "cos_water": cli_cos_water,
                        "cos_hotel": cli_cos_hotel,
                        "total_cos": cli_total_cos,
                        "gp_sale_of_real_estates": cli_gp_sale_of_real_estates,
                        "gp_rental": cli_gp_rental,
                        "gp_management_fees": cli_gp_management_fees,
                        "gp_hotel": cli_gp_hotel,
                        "gp_water_income": cli_gp_water_income,
                        "total_gross_profit": cli_total_gross_profit,
                        "operating_expenses_advertising": cli_operating_expenses_advertising,
                        "operating_expenses_association_dues": cli_operating_expenses_association_dues,
                        "operating_expenses_commissions": cli_operating_expenses_commissions,
                        "operating_expenses_communications": cli_operating_expenses_communications,
                        "operating_expenses_depreciation_and_amortization": cli_operating_expenses_depreciation_and_amortization,
                        "operating_expenses_donations": cli_operating_expenses_donations,
                        "operating_expenses_fuel_and_lubricants": cli_operating_expenses_fuel_and_lubricants,
                        "operating_expenses_impairment_losses": cli_operating_expenses_impairment_losses,
                        "operating_expenses_insurance": cli_operating_expenses_insurance,
                        "operating_expenses_management_fee_expense": cli_operating_expenses_management_fee_expense,
                        "operating_expenses_move_in_fees": cli_operating_expenses_move_in_fees,
                        "operating_expenses_penalties": cli_operating_expenses_penalties,
                        "operating_expenses_professional_and_legal_fees": cli_operating_expenses_professional_and_legal_fees,
                        "operating_expenses_rent": cli_operating_expenses_rent,
                        "operating_expenses_repairs_and_maintenance": cli_operating_expenses_repairs_and_maintenance,
                        "operating_expenses_representation_and_entertainment": cli_operating_expenses_representation_and_entertainment,
                        "operating_expenses_salaries_and_employee_benefits": cli_operating_expenses_salaries_and_employee_benefits,
                        "operating_expenses_security_and_janitorial_services": cli_operating_expenses_security_and_janitorial_services,
                        "operating_expenses_subscription_and_membership_dues": cli_operating_expenses_subscription_and_membership_dues,
                        "operating_expenses_supplies": cli_operating_expenses_supplies,
                        "operating_expenses_taxes_and_licenses": cli_operating_expenses_taxes_and_licenses,
                        "operating_expenses_trainings_and_seminars": cli_operating_expenses_trainings_and_seminars,
                        "operating_expenses_transportation_and_travel": cli_operating_expenses_transportation_and_travel,
                        "operating_expenses_utilities": cli_operating_expenses_utilities,
                        "operating_expenses_other_operating_expenses": cli_operating_expenses_other_operating_expenses,
                        "total_operating_expenses": cli_total_operating_expenses,
                        "reversal_of_payables": cli_reversal_of_payables,
                        "administrative_charges": cli_administrative_charges,
                        "reservation_fees_foregone": cli_reservation_fees_foregone,
                        "late_payment_penalties": cli_late_payment_penalties,
                        "documentation_fees": cli_documentation_fees,
                        "sales_water": cli_sales_water,
                        "water_income": cli_water_income,
                        "referal_incentives": cli_referal_incentives,
                        "unrealized_foreign_exchange_gain": cli_unrealized_foreign_exchange_gain,
                        "realized_foreign_exchange_gain": cli_realized_foreign_exchange_gain,
                        "others": cli_others,
                        "other_operating_income": cli_other_operating_income,
                        "net_operating_income": cli_net_operating_income,
                        "equity_in_net_earnings_losses": cli_equity_in_net_earnings_losses,
                        "equity_in_net_earnings": cli_equity_in_net_earnings,
                        "loss_on_sale_of_asset": cli_loss_on_sale_of_asset,
                        "interest_expense_on_loans": cli_interest_expense_on_loans,
                        "interest_expense_on_defined_benefit_obligation": cli_interest_expense_on_defined_benefit_obligation,
                        "amortized_debt_issuance_cost": cli_amortized_debt_issuance_cost,
                        "expected_credit_losses": cli_expected_credit_losses,
                        "interest_income_from_bank_deposits": cli_interest_income_from_bank_deposits,
                        "bank_charges": cli_bank_charges,
                        "interest_income_from_in_house_financing": cli_interest_income_from_in_house_financing,
                        "gain_on_sale_of_financial_assets": cli_gain_on_sale_of_financial_assets,
                        "other_gains": cli_other_gains,
                        "gain_on_sale_of_property": cli_gain_on_sale_of_property,
                        "realized_foreign_exchange_loss": cli_realized_foreign_exchange_loss,
                        "unrealized_foreign_exchange_loss": cli_unrealized_foreign_exchange_loss,
                        "other_finance_income": cli_other_finance_income,
                        "discount_on_non_current_contract_receivables": cli_discount_on_non_current_contract_receivables,
                        "other_finance_costs": cli_other_finance_costs,
                        "interest_expense_lease_liability": cli_interest_expense_lease_liability,
                        "total_other_income_or_expense": cli_total_other_income_or_expense,
                        "net_profit_before_tax": cli_net_profit_before_tax,
                        "current_income_tax": cli_current_income_tax,
                        "final_income_tax": cli_final_income_tax,
                        "deferred_income_tax": cli_deferred_income_tax,
                        "total_consolidated_net_income": cli_total_consolidated_net_income,
                        "total_nci": cli_total_nci,
                        "total_net_income_attributable_to_parent": cli_total_net_income_attributable_to_parent,
                        "consolidated_niat": cli_consolidated_niat,
                        "parent_niat": cli_parent_niat,
                        "gpm": cli_gpm,
                        "opex_ratio": cli_opex_ratio,
                        "np_margin": cli_np_margin,
                    },
                    "BLCBP": {
                        "revenues_sales_of_real_estates": blcbp_revenues_sales_of_real_estates,
                        "revenues_rental": blcbp_revenues_rental,
                        "revenues_management_fees": blcbp_revenues_management_fees,
                        "revenues_hotel_operations": blcbp_revenues_hotel_operations,
                        "total_revenue": blcbp_total_revenue,
                        "cos_real_estates": blcbp_cos_real_estates,
                        "cos_depreciation": blcbp_cos_depreciation,
                        "cos_taxes": blcbp_cos_taxes,
                        "cos_salaries_and_other_benefits": blcbp_cos_salaries_and_other_benefits,
                        "cos_water": blcbp_cos_water,
                        "cos_hotel": blcbp_cos_hotel,
                        "total_cos": blcbp_total_cos,
                        "gp_sale_of_real_estates": blcbp_gp_sale_of_real_estates,
                        "gp_rental": blcbp_gp_rental,
                        "gp_management_fees": blcbp_gp_management_fees,
                        "gp_hotel": blcbp_gp_hotel,
                        "gp_water_income": blcbp_gp_water_income,
                        "total_gross_profit": blcbp_total_gross_profit,
                        "operating_expenses_advertising": blcbp_operating_expenses_advertising,
                        "operating_expenses_association_dues": blcbp_operating_expenses_association_dues,
                        "operating_expenses_commissions": blcbp_operating_expenses_commissions,
                        "operating_expenses_communications": blcbp_operating_expenses_communications,
                        "operating_expenses_depreciation_and_amortization": blcbp_operating_expenses_depreciation_and_amortization,
                        "operating_expenses_donations": blcbp_operating_expenses_donations,
                        "operating_expenses_fuel_and_lubricants": blcbp_operating_expenses_fuel_and_lubricants,
                        "operating_expenses_impairment_losses": blcbp_operating_expenses_impairment_losses,
                        "operating_expenses_insurance": blcbp_operating_expenses_insurance,
                        "operating_expenses_management_fee_expense": blcbp_operating_expenses_management_fee_expense,
                        "operating_expenses_move_in_fees": blcbp_operating_expenses_move_in_fees,
                        "operating_expenses_penalties": blcbp_operating_expenses_penalties,
                        "operating_expenses_professional_and_legal_fees": blcbp_operating_expenses_professional_and_legal_fees,
                        "operating_expenses_rent": blcbp_operating_expenses_rent,
                        "operating_expenses_repairs_and_maintenance": blcbp_operating_expenses_repairs_and_maintenance,
                        "operating_expenses_representation_and_entertainment": blcbp_operating_expenses_representation_and_entertainment,
                        "operating_expenses_salaries_and_employee_benefits": blcbp_operating_expenses_salaries_and_employee_benefits,
                        "operating_expenses_security_and_janitorial_services": blcbp_operating_expenses_security_and_janitorial_services,
                        "operating_expenses_subscription_and_membership_dues": blcbp_operating_expenses_subscription_and_membership_dues,
                        "operating_expenses_supplies": blcbp_operating_expenses_supplies,
                        "operating_expenses_taxes_and_licenses": blcbp_operating_expenses_taxes_and_licenses,
                        "operating_expenses_trainings_and_seminars": blcbp_operating_expenses_trainings_and_seminars,
                        "operating_expenses_transportation_and_travel": blcbp_operating_expenses_transportation_and_travel,
                        "operating_expenses_utilities": blcbp_operating_expenses_utilities,
                        "operating_expenses_other_operating_expenses": blcbp_operating_expenses_other_operating_expenses,
                        "total_operating_expenses": blcbp_total_operating_expenses,
                        "reversal_of_payables": blcbp_reversal_of_payables,
                        "administrative_charges": blcbp_administrative_charges,
                        "reservation_fees_foregone": blcbp_reservation_fees_foregone,
                        "late_payment_penalties": blcbp_late_payment_penalties,
                        "documentation_fees": blcbp_documentation_fees,
                        "sales_water": blcbp_sales_water,
                        "water_income": blcbp_water_income,
                        "referal_incentives": blcbp_referal_incentives,
                        "unrealized_foreign_exchange_gain": blcbp_unrealized_foreign_exchange_gain,
                        "realized_foreign_exchange_gain": blcbp_realized_foreign_exchange_gain,
                        "others": blcbp_others,
                        "other_operating_income": blcbp_other_operating_income,
                        "net_operating_income": blcbp_net_operating_income,
                        "equity_in_net_earnings_losses": blcbp_equity_in_net_earnings_losses,
                        "equity_in_net_earnings": blcbp_equity_in_net_earnings,
                        "loss_on_sale_of_asset": blcbp_loss_on_sale_of_asset,
                        "interest_expense_on_loans": blcbp_interest_expense_on_loans,
                        "interest_expense_on_defined_benefit_obligation": blcbp_interest_expense_on_defined_benefit_obligation,
                        "amortized_debt_issuance_cost": blcbp_amortized_debt_issuance_cost,
                        "expected_credit_losses": blcbp_expected_credit_losses,
                        "interest_income_from_bank_deposits": blcbp_interest_income_from_bank_deposits,
                        "bank_charges": blcbp_bank_charges,
                        "interest_income_from_in_house_financing": blcbp_interest_income_from_in_house_financing,
                        "gain_on_sale_of_financial_assets": blcbp_gain_on_sale_of_financial_assets,
                        "other_gains": blcbp_other_gains,
                        "gain_on_sale_of_property": blcbp_gain_on_sale_of_property,
                        "realized_foreign_exchange_loss": blcbp_realized_foreign_exchange_loss,
                        "unrealized_foreign_exchange_loss": blcbp_unrealized_foreign_exchange_loss,
                        "other_finance_income": blcbp_other_finance_income,
                        "discount_on_non_current_contract_receivables": blcbp_discount_on_non_current_contract_receivables,
                        "other_finance_costs": blcbp_other_finance_costs,
                        "interest_expense_lease_liability": blcbp_interest_expense_lease_liability,
                        "total_other_income_or_expense": blcbp_total_other_income_or_expense,
                        "net_profit_before_tax": blcbp_net_profit_before_tax,
                        "current_income_tax": blcbp_current_income_tax,
                        "final_income_tax": blcbp_final_income_tax,
                        "deferred_income_tax": blcbp_deferred_income_tax,
                        "total_consolidated_net_income": blcbp_total_consolidated_net_income,
                        "total_nci": blcbp_total_nci,
                        "total_net_income_attributable_to_parent": blcbp_total_net_income_attributable_to_parent,
                        "consolidated_niat": blcbp_consolidated_niat,
                        "parent_niat": blcbp_parent_niat,
                        "gpm": blcbp_gpm,
                        "opex_ratio": blcbp_opex_ratio,
                        "np_margin": blcbp_np_margin,
                    },
                    "YES": {
                        "revenues_sales_of_real_estates": yes_revenues_sales_of_real_estates,
                        "revenues_rental": yes_revenues_rental,
                        "revenues_management_fees": yes_revenues_management_fees,
                        "revenues_hotel_operations": yes_revenues_hotel_operations,
                        "total_revenue": yes_total_revenue,
                        "cos_real_estates": yes_cos_real_estates,
                        "cos_depreciation": yes_cos_depreciation,
                        "cos_taxes": yes_cos_taxes,
                        "cos_salaries_and_other_benefits": yes_cos_salaries_and_other_benefits,
                        "cos_water": yes_cos_water,
                        "cos_hotel": yes_cos_hotel,
                        "total_cos": yes_total_cos,
                        "gp_sale_of_real_estates": yes_gp_sale_of_real_estates,
                        "gp_rental": yes_gp_rental,
                        "gp_management_fees": yes_gp_management_fees,
                        "gp_hotel": yes_gp_hotel,
                        "gp_water_income": yes_gp_water_income,
                        "total_gross_profit": yes_total_gross_profit,
                        "operating_expenses_advertising": yes_operating_expenses_advertising,
                        "operating_expenses_association_dues": yes_operating_expenses_association_dues,
                        "operating_expenses_commissions": yes_operating_expenses_commissions,
                        "operating_expenses_communications": yes_operating_expenses_communications,
                        "operating_expenses_depreciation_and_amortization": yes_operating_expenses_depreciation_and_amortization,
                        "operating_expenses_donations": yes_operating_expenses_donations,
                        "operating_expenses_fuel_and_lubricants": yes_operating_expenses_fuel_and_lubricants,
                        "operating_expenses_impairment_losses": yes_operating_expenses_impairment_losses,
                        "operating_expenses_insurance": yes_operating_expenses_insurance,
                        "operating_expenses_management_fee_expense": yes_operating_expenses_management_fee_expense,
                        "operating_expenses_move_in_fees": yes_operating_expenses_move_in_fees,
                        "operating_expenses_penalties": yes_operating_expenses_penalties,
                        "operating_expenses_professional_and_legal_fees": yes_operating_expenses_professional_and_legal_fees,
                        "operating_expenses_rent": yes_operating_expenses_rent,
                        "operating_expenses_repairs_and_maintenance": yes_operating_expenses_repairs_and_maintenance,
                        "operating_expenses_representation_and_entertainment": yes_operating_expenses_representation_and_entertainment,
                        "operating_expenses_salaries_and_employee_benefits": yes_operating_expenses_salaries_and_employee_benefits,
                        "operating_expenses_security_and_janitorial_services": yes_operating_expenses_security_and_janitorial_services,
                        "operating_expenses_subscription_and_membership_dues": yes_operating_expenses_subscription_and_membership_dues,
                        "operating_expenses_supplies": yes_operating_expenses_supplies,
                        "operating_expenses_taxes_and_licenses": yes_operating_expenses_taxes_and_licenses,
                        "operating_expenses_trainings_and_seminars": yes_operating_expenses_trainings_and_seminars,
                        "operating_expenses_transportation_and_travel": yes_operating_expenses_transportation_and_travel,
                        "operating_expenses_utilities": yes_operating_expenses_utilities,
                        "operating_expenses_other_operating_expenses": yes_operating_expenses_other_operating_expenses,
                        "total_operating_expenses": yes_total_operating_expenses,
                        "reversal_of_payables": yes_reversal_of_payables,
                        "administrative_charges": yes_administrative_charges,
                        "reservation_fees_foregone": yes_reservation_fees_foregone,
                        "late_payment_penalties": yes_late_payment_penalties,
                        "documentation_fees": yes_documentation_fees,
                        "sales_water": yes_sales_water,
                        "water_income": yes_water_income,
                        "referal_incentives": yes_referal_incentives,
                        "unrealized_foreign_exchange_gain": yes_unrealized_foreign_exchange_gain,
                        "realized_foreign_exchange_gain": yes_realized_foreign_exchange_gain,
                        "others": yes_others,
                        "other_operating_income": yes_other_operating_income,
                        "net_operating_income": yes_net_operating_income,
                        "equity_in_net_earnings_losses": yes_equity_in_net_earnings_losses,
                        "equity_in_net_earnings": yes_equity_in_net_earnings,
                        "loss_on_sale_of_asset": yes_loss_on_sale_of_asset,
                        "interest_expense_on_loans": yes_interest_expense_on_loans,
                        "interest_expense_on_defined_benefit_obligation": yes_interest_expense_on_defined_benefit_obligation,
                        "amortized_debt_issuance_cost": yes_amortized_debt_issuance_cost,
                        "expected_credit_losses": yes_expected_credit_losses,
                        "interest_income_from_bank_deposits": yes_interest_income_from_bank_deposits,
                        "bank_charges": yes_bank_charges,
                        "interest_income_from_in_house_financing": yes_interest_income_from_in_house_financing,
                        "gain_on_sale_of_financial_assets": yes_gain_on_sale_of_financial_assets,
                        "other_gains": yes_other_gains,
                        "gain_on_sale_of_property": yes_gain_on_sale_of_property,
                        "realized_foreign_exchange_loss": yes_realized_foreign_exchange_loss,
                        "unrealized_foreign_exchange_loss": yes_unrealized_foreign_exchange_loss,
                        "other_finance_income": yes_other_finance_income,
                        "discount_on_non_current_contract_receivables": yes_discount_on_non_current_contract_receivables,
                        "other_finance_costs": yes_other_finance_costs,
                        "interest_expense_lease_liability": yes_interest_expense_lease_liability,
                        "total_other_income_or_expense": yes_total_other_income_or_expense,
                        "net_profit_before_tax": yes_net_profit_before_tax,
                        "current_income_tax": yes_current_income_tax,
                        "final_income_tax": yes_final_income_tax,
                        "deferred_income_tax": yes_deferred_income_tax,
                        "total_consolidated_net_income": yes_total_consolidated_net_income,
                        "total_nci": yes_total_nci,
                        "total_net_income_attributable_to_parent": yes_total_net_income_attributable_to_parent,
                        "consolidated_niat": yes_consolidated_niat,
                        "parent_niat": yes_parent_niat,
                        "gpm": yes_gpm,
                        "opex_ratio": yes_opex_ratio,
                        "np_margin": yes_np_margin,
                    },
                }
            ),
            200,
        )
    else:
        return jsonify({"error": "Invalid file type"}), 400


if __name__ == "__main__":
    app.run(debug=True)
