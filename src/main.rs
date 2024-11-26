use clap::{Arg, Command};
use calamine::{open_workbook, Reader, Xlsx, DataType};
use chrono::{Datelike, NaiveDate, Utc};
use std::{collections::HashMap, error::Error, fs::File, io::{BufRead, BufReader, BufWriter, Write}};

#[derive(Debug)]
struct Employee {
    emp_id: i32,
    emp_name: String,
    dept_id: i32,
    mobile_no: String,
    email: String,
}

#[derive(Debug)]
struct Department {
    dept_id: i32,
    dept_title: String,
}

fn main() -> Result<(), Box<dyn Error>> {

    let matches = Command::new("Employee Report Generator")
        .arg(
            Arg::new("emp_data")
                .short('e')
                .long("emp-data-file-path")
                .required(true)
                .value_parser(clap::value_parser!(String))
                .help("Path to the employee data file"),
        )
        .arg(
            Arg::new("dept_data")
                .short('d')
                .long("dept-data-file-path")
                .required(true)
                .value_parser(clap::value_parser!(String))
                .help("Path to the department data file"),
        )
        .arg(
            Arg::new("salary_data")
                .short('s')
                .long("salary-data-file-path")
                .required(true)
                .value_parser(clap::value_parser!(String))
                .help("Path to the salary data file"),
        )
        .arg(
            Arg::new("leave_data")
                .short('l')
                .long("leave-data-file-path")
                .required(true)
                .value_parser(clap::value_parser!(String))
                .help("Path to the leave data file"),
        )
        .arg(
            Arg::new("output_file")
                .short('o')
                .long("output-file-path")
                .required(true)
                .value_parser(clap::value_parser!(String))
                .help("Path to the output file"),
        )
        .get_matches();

    //Storing file path
    let default_emp_data_path = "default_emp_data_path".to_string();
    let default_dept_data_path = "default_dept_data_path".to_string();
    let default_salary_data_path = "default_salary_data_path".to_string();
    let default_leave_data_path = "default_leave_data_path".to_string();
    let default_output_path = "default_output_file_path".to_string();
    let emp_data_file_path = matches.get_one::<String>("emp_data").unwrap_or(&default_emp_data_path);
    let dept_data_file_path = matches.get_one::<String>("dept_data").unwrap_or(&default_dept_data_path);
    let salary_data_file_path = matches.get_one::<String>("salary_data").unwrap_or(&default_salary_data_path);
    let leave_data_file_path = matches.get_one::<String>("leave_data").unwrap_or(&default_leave_data_path);
    let output_file_path = matches.get_one::<String>("output_file").unwrap_or(&default_output_path);

    // Parsing files
    let employees = parse_employee_data(emp_data_file_path).expect("Parsing error");
    let departments = parse_department_data(dept_data_file_path).expect("Parsing error");
    let salaries = parse_salary_data(salary_data_file_path).expect("Parsing error");
    let leaves = parse_leave_data(leave_data_file_path).expect("Parsing error");


    // Generating the output
    generate_output(employees, departments, salaries, leaves, output_file_path).expect("output gen error");
    Ok(())
}

fn parse_employee_data(file_path: &str) -> Result<Vec<Employee>, Box<dyn Error>> {
    let file = File::open(file_path).expect("error in opening file");
    let reader = BufReader::new(file);
    let mut lines = reader.lines();
    lines.next(); // Skip header
    let mut employees = Vec::new();
    for line in lines {
        let line = line.expect("Line was not loaded");
        let fields: Vec<&str> = line.split('|').collect();
        let employee = Employee {
            emp_id: fields[0].parse().unwrap_or(-1),
            emp_name: fields[1].to_string(),
            dept_id: fields[2].parse().unwrap_or(-1),
            mobile_no: fields[3].to_string(),
            email: fields[4].to_string(),
        };
        employees.push(employee);
    }

    Ok(employees)
}

fn parse_department_data(file_path: &str) -> Result<HashMap<i32, Department>, Box<dyn Error>> {
    let mut workbook: Xlsx<_> = open_workbook(file_path).expect("error in opening workbook");
    let sheet = workbook.worksheet_range("Sheet1").expect("error in opening sheet");
    let mut departments = HashMap::new();

    for row in sheet.rows().skip(1) {
        let department = Department {
            dept_id: row[0].get_float().unwrap_or(-1.1) as i32,
            dept_title: row[1].get_string().unwrap_or("no string found").to_string(),
        };
        departments.insert(department.dept_id, department);
    }

    Ok(departments)
}

fn parse_salary_data(file_path: &str) -> Result<HashMap<i32, String>, Box<dyn Error>> {
    let mut workbook: Xlsx<_> = open_workbook(file_path).expect("error in opening workbook");
    let sheet = workbook.worksheet_range("Sheet1").expect("error in loading sheet");
    let mut salaries = HashMap::new();
    let current_month = Utc::now().month();
    for row in sheet.rows().skip(1) {
        let emp_id = row[0].get_float().unwrap_or(-1.1) as i32;
        let salary_date = row[2].get_string().unwrap_or("default date").to_string();
        let formatted_date=format!("01 {}",salary_date);

        let parsed_date = NaiveDate::parse_from_str(&formatted_date, "%d %b %Y").expect("error in parsing formatted date");
        let salary_month = parsed_date.month();
    
        if salary_month == current_month {
            salaries.insert(emp_id, row[4].get_string().unwrap_or("default date").to_string());
        }
    }

    Ok(salaries)
}

fn parse_leave_data(file_path: &str) -> Result<HashMap<i32, i32>, Box<dyn Error>> {
    let mut workbook: Xlsx<_> = open_workbook(file_path).expect("error in opening workbook");
    let sheet = workbook.worksheet_range("Sheet1").expect("error in loading sheet");
    let mut leaves = HashMap::new();
    let current_month = Utc::now().month();

    for row in sheet.rows().skip(1) {
        let emp_id = row[0].get_float().unwrap_or(-1.1) as i32;
        let leave_from = NaiveDate::parse_from_str(row[2].get_string().expect("error in parsing from date"), "%d-%m-%Y").expect("error in parsing from date");
        let leave_to = NaiveDate::parse_from_str(row[3].get_string().expect("error in parsing from date"), "%d-%m-%Y").expect("error in parsing to date");
        if leave_from.month()==current_month && leave_to.month()==current_month{
            let days = (leave_to - leave_from).num_days() as i32 + 1;
            *leaves.entry(emp_id).or_insert(0) += days;
        }else{
            if leave_from.month() == current_month || leave_to.month() == current_month {
                let mut tochange=0;
                let from = if leave_from.month() == current_month {
                    leave_from
                } else {
                    NaiveDate::from_ymd_opt(leave_from.year(), current_month, 1).expect("wrong date")
                };
    
                let to = if leave_to.month() == current_month {
                    leave_to
                } else {
                    tochange=1;
                    NaiveDate::from_ymd_opt(leave_to.year(), current_month + 1, 1).expect("wrong date")

                };
    
                let mut days = (to - from).num_days() as i32;
                if tochange==0{
                    days+=1;
                }
                *leaves.entry(emp_id).or_insert(0) += days;
            }else if leave_from.month()<current_month && leave_to.month()>current_month{
                let to=NaiveDate::from_ymd_opt(leave_to.year(),current_month+1,1).expect("wrong date");
                let from=NaiveDate::from_ymd_opt(leave_to.year(),current_month,1).expect("wrong date");
                let days=(to-from).num_days() as i32;
                *leaves.entry(emp_id).or_insert(0)+=days;
            }
        }
        
    }

    Ok(leaves)
}

fn generate_output(
    employees: Vec<Employee>,
    departments: HashMap<i32, Department>,
    salaries: HashMap<i32, String>,
    leaves: HashMap<i32, i32>,
    output_file_path: &str,
) -> Result<(), Box<dyn Error>> {
    let mut writer = BufWriter::new(File::create(output_file_path).expect("Output file path not resolved"));

    //header
    writeln!(
        writer,
        "Emp ID~#~Emp Name~#~Dept Title~#~Mobile No~#~Email~#~Salary Status~#~On Leave"
    ).expect("header writing error");

    for employee in employees {
        let dept_title = departments
        .get(&employee.dept_id)
        .map(|d| d.dept_title.clone())
        .unwrap_or_else(|| "Unknown".to_string());
        let salary_status = salaries
        .get(&employee.emp_id)
        .cloned() 
        .unwrap_or_else(|| "Not Credited".to_string());


        let leave_days = leaves.get(&employee.emp_id).unwrap_or(&0);

        writeln!(
            writer,
            "{}~#~{}~#~{}~#~{}~#~{}~#~{}~#~{}",
            employee.emp_id,
            employee.emp_name,
            dept_title,
            employee.mobile_no,
            employee.email,
            salary_status,
            leave_days
        ).expect("output record error");
    }

    Ok(())
}