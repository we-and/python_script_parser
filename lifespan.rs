use chrono::{Duration, Local, NaiveDate, NaiveTime};
use rand::Rng;
use std::env;
use std::process;

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() != 4 {
        eprintln!("Usage: {} <app_name> <date> <number>", args[0]);
        process::exit(1);
    }

    let app_name = &args[1];
    let input_date = NaiveDate::parse_from_str(&args[2], "%m/%d/%Y").unwrap();
    let input_number = args[3].parse::<i32>().unwrap();

    let now = Local::now().naive_local();

    let date1 = input_date.and_time(NaiveTime::from_hms(0, 0, 0)).unwrap()
        + Duration::seconds(rand::thread_rng().gen_range(-1200..=1200));
    if now > date1 {
        process::exit(0);
    }

    let date2 = input_date.and_time(NaiveTime::from_hms(0, 0, 0)).unwrap()
        + Duration::seconds(rand::thread_rng().gen_range(-1200..=1200));
    if now > date2 {
        process::exit(0);
    }

    // ...

    let date10 = input_date.and_time(NaiveTime::from_hms(0, 0, 0)).unwrap()
        + Duration::seconds(876);
    if now > date10 {
        process::exit(0);
    }

    println!("No conditions met. Continuing execution...");
}