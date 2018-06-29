package de.morphbit.poi.model;

import java.util.function.Predicate;

public class ExcelDataTemplateOptions {

	private int ignoreFirstLinesCount;
	private Predicate<?> filter;
	
	public ExcelDataTemplateOptions() {
		
	}
	
	public ExcelDataTemplateOptions(int ignoreFirstLinesCount, Predicate<?> filter) {
		this.ignoreFirstLinesCount = ignoreFirstLinesCount;
		this.filter = filter;
	}

	public int getIgnoreFirstLinesCount() {
		return ignoreFirstLinesCount;
	}
	public void setIgnoreFirstLinesCount(int ignoreFirstLinesCount) {
		this.ignoreFirstLinesCount = ignoreFirstLinesCount < 0 ? 0 : ignoreFirstLinesCount;
	}
	
	public Predicate<?> getFilter() {
		return filter;
	}
	public void setFilter(Predicate<?> filter) {
		this.filter = filter;
	}
	
	public ExcelDataTemplateOptions withIgnoreFirstLinesCount(int ignoreFirstLinesCount) {
		setIgnoreFirstLinesCount(ignoreFirstLinesCount);
		return this;
	}
	
	public <T> ExcelDataTemplateOptions withFilter(Predicate<? super T> filter) {
		setFilter(filter);
		return this;
	}
}
