require "application_system_test_case"

class ConvertsTest < ApplicationSystemTestCase
  setup do
    @convert = converts(:one)
  end

  test "visiting the index" do
    visit converts_url
    assert_selector "h1", text: "Converts"
  end

  test "creating a Convert" do
    visit converts_url
    click_on "New Convert"

    fill_in "File", with: @convert.file
    click_on "Create Convert"

    assert_text "Convert was successfully created"
    click_on "Back"
  end

  test "updating a Convert" do
    visit converts_url
    click_on "Edit", match: :first

    fill_in "File", with: @convert.file
    click_on "Update Convert"

    assert_text "Convert was successfully updated"
    click_on "Back"
  end

  test "destroying a Convert" do
    visit converts_url
    page.accept_confirm do
      click_on "Destroy", match: :first
    end

    assert_text "Convert was successfully destroyed"
  end
end
