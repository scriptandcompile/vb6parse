// This whole file is just a horrible hack to allow grouping of the CST parsing tests into
// a separate folder without needing to throw all the tests into a single huge file or
// throwing all test files in the root tests folder.

mod assignment;
mod call;
mod chdir_statement;
mod chdrive_statement;
mod comments;
mod cst;
mod cst_navigation;
mod declaration;
mod do_loop;
mod do_loop_in_function;
mod exit_statement;
mod for_loop;
mod function_linecont;
mod goto_statement;
mod label;
mod option_explicit;
mod property_statement;
mod redim_statement;
mod select_case;
mod set_statement;
mod sub;
mod with_statement;
