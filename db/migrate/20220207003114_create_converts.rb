class CreateConverts < ActiveRecord::Migration[6.1]
  def change
    create_table :converts do |t|
      t.string :file

      t.timestamps
    end
  end
end
