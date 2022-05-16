require "test_helper"

class ConvertsControllerTest < ActionDispatch::IntegrationTest
  setup do
    @convert = converts(:one)
  end

  test "should get index" do
    get converts_url
    assert_response :success
  end

  test "should get new" do
    get new_convert_url
    assert_response :success
  end

  test "should create convert" do
    assert_difference('Convert.count') do
      post converts_url, params: { convert: { file: @convert.file } }
    end

    assert_redirected_to convert_url(Convert.last)
  end

  test "should show convert" do
    get convert_url(@convert)
    assert_response :success
  end

  test "should get edit" do
    get edit_convert_url(@convert)
    assert_response :success
  end

  test "should update convert" do
    patch convert_url(@convert), params: { convert: { file: @convert.file } }
    assert_redirected_to convert_url(@convert)
  end

  test "should destroy convert" do
    assert_difference('Convert.count', -1) do
      delete convert_url(@convert)
    end

    assert_redirected_to converts_url
  end
end
